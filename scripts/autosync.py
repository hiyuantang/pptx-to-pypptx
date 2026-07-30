#!/usr/bin/env python3
"""Auto-sync deck code from a PowerPoint edit -- the deck->code half of the round trip.

The agent runs this at the start of a deck task (see SKILL.md). It cheaply checks
whether ``out/<name>.pptx`` changed since the last sync (a build or a previous
auto-sync). If a human edited the deck in PowerPoint, it regenerates only the
affected ``slides/*.py`` so the code matches the deck.

Deliberately narrow:
  * It syncs **deck -> code only**; it never rebuilds (that would overwrite the
    file you just saved and could conflict with PowerPoint holding it open).
  * It never fails the caller -- any error is reported and swallowed (exit 0), so
    a sync hiccup can't derail the task.
  * It does no TODO review or verification -- that stays a human/agent decision.

Thin orchestrator: change detection and stamping live in the project's
``lib/roundtrip_state.py`` (shared with ``build_deck.py``); slide code generation
reuses ``generate_slides.py``; project discovery reuses ``detect_project.py``.

Usage:
    uv run --directory <project> python <skill>/scripts/autosync.py --project-dir <project>
"""

import argparse
import contextlib
import importlib.util
import io
import re
import sys
from pathlib import Path

_SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(_SCRIPTS))

# Light imports only (stdlib) -- safe even before the project's deps are installed.
from detect_project import _find_projects, _output_pptx  # noqa: E402


def _load_state_module(project_dir: Path):
    """Load the project's own ``lib/roundtrip_state.py`` by path.

    Loading by file path (rather than via ``sys.path``) keeps side-by-side
    projects from colliding on the ``lib`` package name, and guarantees autosync
    uses the exact same hashing logic ``build_deck.py`` stamped with.
    """
    path = project_dir / "lib" / "roundtrip_state.py"
    if not path.is_file():
        return None
    spec = importlib.util.spec_from_file_location(
        f"_roundtrip_state_{abs(hash(str(project_dir)))}", path
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def sync_project(project_dir: Path) -> str:
    """Sync one project's code from its edited deck.

    Always returns ONE decisive status line for the agent to act on without any
    further checking:
      * ``OK — …``      nothing to do; code already matches the deck. Proceed.
      * ``SYNCED — …``  code was regenerated to match the deck. Proceed.
      * ``SKIPPED — …`` could not sync; the message says what to do.
    """
    name = project_dir.name
    state_mod = _load_state_module(project_dir)
    if state_mod is None:
        return (f"{name}: SKIPPED — predates auto-sync (no lib/roundtrip_state.py). "
                f"Re-scaffold to enable; deck code may be stale.")

    _, out_pptx = _output_pptx(project_dir)
    if out_pptx is None or not out_pptx.exists():
        return f"{name}: OK — deck not built yet; nothing to sync. Proceed."

    new_state = state_mod.compute_state(out_pptx)
    old_state = state_mod.read_state(project_dir)

    # Slides first: a regenerated slide file already gets its comment region from
    # the deck, so those slides need no second pass.
    status, regenerated = _sync_slides(
        project_dir, state_mod, out_pptx, name, old_state, new_state
    )

    # Comments live in their own part, so a reviewer replying in PowerPoint leaves
    # every slide hash untouched and the code above reports "no changes". Patch the
    # comment region of those slide files directly, leaving the rest of each file
    # byte-identical.
    comment_note, comments_ok = _sync_comments(
        project_dir, out_pptx, state_mod, old_state, new_state, regenerated
    )

    # The marker is advanced only once BOTH halves are done. Writing it earlier
    # would mark a comment change as handled even if its patch failed, and the
    # next run would compare against the new digest and never retry.
    if comments_ok:
        state_mod.write_state(project_dir, new_state)

    if comment_note:
        status = f"{status} ({comment_note})"
    return status


def _slide_file(project_dir: Path, idx: int):
    """The project's slide file for a 1-based deck position, or ``None``."""
    matches = sorted((project_dir / "slides").glob(f"s{idx:02d}_*.py"))
    return matches[0] if matches else None


def _sync_comments(
    project_dir: Path, out_pptx: Path, state_mod, old_state, new_state, regenerated: set
) -> tuple:
    """Patch the comment region of slide files whose comments changed in the deck.

    Makes comments behave like normal PowerPoint comments across the round trip:
    whatever the human left in the deck (replies added, comments deleted, threads
    resolved) becomes what the slide file says, so the next build reproduces
    exactly that and never resurrects a removed comment.

    Only the fenced region is rewritten -- hand-written slide code in the same file
    is untouched. Returns ``(note, ok)``: a short note when something changed, and
    ``ok=False`` if any slide could not be patched, which keeps the caller from
    advancing the round-trip marker so the next run retries. Never raises: a
    comment hiccup must not derail the sync.
    """
    try:
        from helpers.comments import (
            extract_authors,
            extract_comments,
            merge_authors_json,
            parse_comment_calls,
            patch_comment_region,
            render_comment_calls,
        )
    except Exception:
        return "", True
    try:
        targets = set(state_mod.changed_comment_slides(old_state, new_state))

        # Slides carrying an unbuilt (pending) comment are always re-checked, even
        # with an unchanged digest: build_deck.py stamps a fresh marker, so nothing
        # else would ever revisit them. Once the deck has the comment the deck copy
        # wins and the flag clears itself; while it is still unbuilt the render is
        # identical and the patch is a no-op. Leaving the flag set would be worse
        # than useless -- a later deletion in PowerPoint would be carried over and
        # the comment resurrected.
        for path in sorted((project_dir / "slides").glob("s*.py")):
            m = re.match(r"s(\d+)_", path.name)
            if m and "pending=True" in path.read_text(encoding="utf-8"):
                targets.add(int(m.group(1)))

        targets = sorted(n for n in targets if n not in regenerated)
        if not targets:
            return "", True

        deck_comments = extract_comments(out_pptx)
        merge_authors_json(extract_authors(out_pptx), project_dir / "lib")

        patched, failed = [], []
        for idx in targets:
            path = _slide_file(project_dir, idx)
            if path is None:
                continue
            source = path.read_text(encoding="utf-8")
            threads = list(deck_comments.get(idx) or [])

            # A comment added by add_comment.py but not built yet exists only in
            # the code. A blind mirror cannot tell that apart from a comment the
            # human deleted in PowerPoint, so carry pending ones across; the flag
            # clears itself once the deck has the comment (the deck copy wins).
            built = {(t.get("author"), t.get("created")) for t in threads}
            threads += [
                t
                for t in parse_comment_calls(source)
                if t.get("pending") and (t.get("author"), t.get("created")) not in built
            ]

            try:
                new_source = patch_comment_region(source, render_comment_calls(threads))
            except ValueError:
                failed.append(idx)  # keep the old file; never trade a slide for a comment
                continue
            if new_source is not None:
                path.write_text(new_source, encoding="utf-8")
                patched.append(idx)

        if not patched and not failed:
            return "", True
        note = ""
        if patched:
            note = f"comments synced on slide(s) [{', '.join(map(str, patched))}]"
        if failed:
            note = (note + "; " if note else "") + (
                f"comment patch FAILED on slide(s) [{', '.join(map(str, failed))}] — will retry"
            )
        return note, not failed
    except Exception as exc:
        # Report rather than swallow: a silent return here would let the caller
        # advance the marker and lose the change.
        return f"comment sync skipped ({type(exc).__name__})", False


def _sync_slides(
    project_dir: Path, state_mod, out_pptx: Path, name: str, old_state, new_state
) -> tuple:
    """Regenerate ``slides/*.py`` from the deck when its slides changed.

    Returns ``(status_line, regenerated_slide_numbers)``; the caller skips the
    comment pass for regenerated slides, whose files already carry the deck's
    comments straight from codegen. The caller also owns writing the round-trip
    marker, so a failed comment patch cannot be recorded as handled.
    """
    # No baseline yet -> establish one without regenerating (assume in sync).
    if old_state is None:
        return f"{name}: OK — baseline recorded; code matches the deck. Proceed.", set()

    # Fast, authoritative gate: identical per-slide hashes -> no slide to rebuild.
    # Comments are hashed separately, so this is not the end of the story.
    if old_state.get("slides") == new_state.get("slides"):
        return f"{name}: OK — no changes; code matches the deck. Proceed.", set()

    changed = state_mod.changed_slides(old_state, new_state)
    total = new_state.get("slide_count", 0)
    old_total = old_state.get("slide_count", 0)

    # Heavy import only once we know we must regenerate.
    from generate_slides import generate_slides

    if total != old_total:
        # Structural change (add/delete/reorder): regenerate every slide so
        # slides/*.py matches the deck exactly. Only files past the new end are
        # deleted up front -- generate_slides renames the rest itself, and deleting
        # them here would discard any unbuilt (pending) comment they carry.
        for stale in (project_dir / "slides").glob("s*.py"):
            m = re.match(r"s(\d+)_", stale.name)
            if m and int(m.group(1)) > total:
                stale.unlink()
        target_slides = list(range(1, total + 1))
        detail = f"deck changed {old_total} -> {total} slides; regenerated all {total}"
    else:
        target_slides = changed
        detail = f"{len(changed)} slide(s) [{', '.join(map(str, changed))}] from PowerPoint edits"

    # Suppress generate_slides' own stdout so only the concise status is printed.
    with contextlib.redirect_stdout(io.StringIO()):
        generate_slides(out_pptx, project_dir, target_slides)

    return (
        f"{name}: SYNCED — {detail}; code now matches the deck. Proceed.",
        set(target_slides),
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Auto-sync deck code from PowerPoint edits")
    parser.add_argument(
        "--project-dir",
        default=".",
        help="Project directory, or a parent to scan (default: current dir)",
    )
    args = parser.parse_args()

    start = Path(args.project_dir).resolve()
    lines = []
    try:
        projects = _find_projects(start)
        if not projects:
            print(f"autosync: OK — no project found under {start}; nothing to sync. Proceed.")
            return
        for project in projects:
            try:
                lines.append(sync_project(project))
            except Exception as exc:  # a broken deck must never derail the task
                lines.append(f"{project.name}: SKIPPED — {type(exc).__name__}: {exc}; "
                             f"deck code may be stale.")
    except Exception as exc:
        lines.append(f"SKIPPED — {type(exc).__name__}: {exc}")

    # Always print exactly one clear status per project; never silent.
    for line in lines:
        print(f"autosync: {line}")
    # Always exit 0 so a sync hiccup can't derail the caller's task.


if __name__ == "__main__":
    main()
