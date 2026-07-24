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
import hashlib
import importlib.util
import io
import json
import re
import shutil
import sys
import uuid
import zipfile
from pathlib import Path

# Comments added by add_comment.py that are not yet in the built deck.
_PENDING_FILE = ".pending.json"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"

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

    # Comments round-trip like slides: mirror the deck's current comments back
    # into the store first, so a human reply/deletion in PowerPoint sticks and is
    # not resurrected by the next build. Runs on every path (a human may edit
    # only comments, leaving slide hashes unchanged).
    comment_note = _sync_comments(project_dir, out_pptx)

    status = _sync_slides(project_dir, state_mod, out_pptx, name)
    if comment_note:
        status = f"{status} ({comment_note})"
    return status


def _comments_signature(comments_dir: Path):
    """Content hash of the comment store, or ``None`` if there is no store."""
    if not comments_dir.is_dir():
        return None
    h = hashlib.sha256()
    for path in sorted(comments_dir.rglob("*")):
        if path.is_file():
            h.update(path.relative_to(comments_dir).as_posix().encode("utf-8"))
            h.update(b"\0")
            h.update(path.read_bytes())
            h.update(b"\0")
    return h.hexdigest()


def _sync_comments(project_dir: Path, out_pptx: Path) -> str:
    """Mirror the built deck's comments back into ``comments/`` (deck -> store).

    Makes modern comments behave like normal PowerPoint comments across the round
    trip: whatever the human left in the deck (replies added, comments deleted)
    becomes the store, so ``build_deck.py`` re-attaches exactly that and never
    brings a removed comment back. Returns a short note when the store changed,
    else "". Never raises — a comment hiccup must not derail the sync.
    """
    try:
        from helpers.comments import extract_comments
    except Exception:
        return ""
    comments_dir = project_dir / "comments"
    carried = []
    try:
        before = _comments_signature(comments_dir)

        # A comment written by add_comment.py but not yet built exists ONLY in
        # the store. A blind mirror cannot tell that apart from a comment the
        # human deleted in PowerPoint, so carry such comments across the mirror
        # (and drop ids that already reached the deck -- those are built now).
        deck_ids = _deck_cm_ids(out_pptx)
        pending_ids = [i for i in _read_pending(comments_dir) if i not in deck_ids]
        carried = _collect_pending(comments_dir, set(pending_ids))
        authors_path = comments_dir / "authors.xml"
        authors_xml = (
            authors_path.read_text(encoding="utf-8") if authors_path.is_file() else None
        )

        # Full re-capture from the deck; rewrites the store when comments exist.
        found = extract_comments(out_pptx, project_dir)
        if found == 0 and comments_dir.is_dir():
            # The deck now has zero comments (e.g. the last one was deleted in
            # PowerPoint). extract_comments leaves the store untouched when it
            # finds none, so drop it here or those comments would re-attach.
            shutil.rmtree(comments_dir)

        _reinject_pending(comments_dir, carried, authors_xml)

        pending_path = comments_dir / _PENDING_FILE
        if pending_ids and comments_dir.is_dir():
            pending_path.write_text(
                json.dumps({"ids": pending_ids}, indent=2), encoding="utf-8"
            )
        elif pending_path.is_file():
            pending_path.unlink()

        after = _comments_signature(comments_dir)
    except Exception:
        return ""
    if before == after:
        return ""
    note = "comments mirrored from deck"
    if carried:
        note += f"; {len(carried)} unbuilt comment(s) preserved"
    return note


def _cm_ids(xml: str) -> set:
    return set(re.findall(r'<p188:cm\b[^>]*?\bid="([^"]+)"', xml))


def _deck_cm_ids(out_pptx: Path) -> set:
    """Comment ids present in the built deck."""
    ids: set = set()
    try:
        with zipfile.ZipFile(out_pptx) as zf:
            for name in zf.namelist():
                if re.match(r"ppt/comments/.*\.xml$", name):
                    ids |= _cm_ids(zf.read(name).decode("utf-8"))
    except Exception:
        pass
    return ids


def _read_pending(comments_dir: Path) -> list:
    path = comments_dir / _PENDING_FILE
    if not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return [i for i in data.get("ids", []) if isinstance(i, str)]
    except Exception:
        return []


def _collect_pending(comments_dir: Path, ids: set) -> list:
    """``[(slide_key, cm_xml)]`` for the given comment ids, read from the store."""
    found = []
    manifest_path = comments_dir / "manifest.json"
    if not ids or not manifest_path.is_file():
        return found
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return found
    for slide, files in (manifest.get("slides") or {}).items():
        for fname in files:
            path = comments_dir / fname
            if not path.is_file():
                continue
            xml = path.read_text(encoding="utf-8")
            for match in re.finditer(r"<p188:cm\b.*?</p188:cm>", xml, re.S):
                frag = match.group(0)
                cm_id = re.search(r'\bid="([^"]+)"', frag)
                if cm_id and cm_id.group(1) in ids:
                    found.append((str(slide), frag))
    return found


def _reinject_pending(comments_dir: Path, pending: list, authors_xml) -> None:
    """Put carried-over comments back after the mirror, one part per slide."""
    if not pending:
        return
    comments_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = comments_dir / "manifest.json"
    try:
        manifest = (
            json.loads(manifest_path.read_text(encoding="utf-8"))
            if manifest_path.is_file()
            else {}
        )
    except Exception:
        manifest = {}
    manifest.setdefault("slides", {})
    manifest.setdefault("authors", "authors.xml")

    # The mirror may have removed authors.xml (deck had no comments); the carried
    # comments still reference their author, so restore it.
    authors_path = comments_dir / manifest["authors"]
    if authors_xml and not authors_path.is_file():
        authors_path.write_text(authors_xml, encoding="utf-8")

    for slide, frag in pending:
        files = manifest["slides"].get(slide) or []
        target = comments_dir / files[0] if files else None
        if target is not None and target.is_file():
            xml = target.read_text(encoding="utf-8")
            target.write_text(
                xml.replace("</p188:cmLst>", frag + "</p188:cmLst>", 1), encoding="utf-8"
            )
            continue
        fname = f"claudeComment_{uuid.uuid4().hex[:12]}.xml"
        (comments_dir / fname).write_text(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p188:cmLst xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p188="{_NS_P188}">'
            f"{frag}</p188:cmLst>",
            encoding="utf-8",
        )
        manifest["slides"][slide] = [fname]
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")


def _sync_slides(project_dir: Path, state_mod, out_pptx: Path, name: str) -> str:
    """Regenerate ``slides/*.py`` from the deck when its slides changed."""
    new_state = state_mod.compute_state(out_pptx)
    old_state = state_mod.read_state(project_dir)

    # No baseline yet -> establish one without regenerating (assume in sync).
    if old_state is None:
        state_mod.write_state(project_dir, new_state)
        return f"{name}: OK — baseline recorded; code matches the deck. Proceed."

    # Fast, authoritative gate: identical per-slide hashes -> nothing to do.
    if old_state.get("slides") == new_state.get("slides"):
        if old_state != new_state:  # size drifted but content identical; refresh.
            state_mod.write_state(project_dir, new_state)
        return f"{name}: OK — no changes; code matches the deck. Proceed."

    changed = state_mod.changed_slides(old_state, new_state)
    total = new_state.get("slide_count", 0)
    old_total = old_state.get("slide_count", 0)

    # Heavy import only once we know we must regenerate.
    from generate_slides import generate_slides

    if total != old_total:
        # Structural change (add/delete/reorder): rebuild the slide files from
        # scratch so slides/*.py matches the deck exactly, then regenerate all.
        for stale in (project_dir / "slides").glob("s*.py"):
            stale.unlink()
        target_slides = list(range(1, total + 1))
        detail = f"deck changed {old_total} -> {total} slides; regenerated all {total}"
    else:
        target_slides = changed
        detail = f"{len(changed)} slide(s) [{', '.join(map(str, changed))}] from PowerPoint edits"

    # Suppress generate_slides' own stdout so only the concise status is printed.
    with contextlib.redirect_stdout(io.StringIO()):
        generate_slides(out_pptx, project_dir, target_slides)

    state_mod.write_state(project_dir, new_state)
    return f"{name}: SYNCED — {detail}; code now matches the deck. Proceed."


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
