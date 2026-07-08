#!/usr/bin/env python3
"""Auto-sync deck code from a PowerPoint edit -- the hook half of the round trip.

``scaffold.py`` registers this as a Claude Code ``UserPromptSubmit`` /
``SessionStart`` hook in the project's ``.claude/settings.local.json``. On each
prompt it cheaply checks whether ``out/<name>.pptx`` changed since the last sync
(a build or a previous auto-sync). If a human edited the deck in PowerPoint, it
regenerates only the affected ``slides/*.py`` so the code matches the deck.

Deliberately narrow:
  * It syncs **deck -> code only**; it never rebuilds (that would overwrite the
    file you just saved and could conflict with PowerPoint holding it open).
  * It never blocks the prompt -- any error is reported and swallowed.
  * It does no TODO review or verification -- that stays a human/agent decision.

Thin orchestrator: change detection and stamping live in the project's
``lib/roundtrip_state.py`` (shared with ``build_deck.py``); slide code generation
reuses ``generate_slides.py``; project discovery reuses ``detect_project.py``.

Run manually:
    uv run --directory <project> python <skill>/scripts/autosync.py --project-dir <project>
"""

import argparse
import contextlib
import importlib.util
import io
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


def sync_project(project_dir: Path) -> str | None:
    """Sync one project's code from its edited deck. Returns a summary, or None
    when nothing needed to happen (the common, near-instant case)."""
    state_mod = _load_state_module(project_dir)
    if state_mod is None:
        return None  # older project without the helper -> skip silently

    _, out_pptx = _output_pptx(project_dir)
    if out_pptx is None or not out_pptx.exists():
        return None  # never built yet -> nothing to sync from

    new_state = state_mod.compute_state(out_pptx)
    old_state = state_mod.read_state(project_dir)

    # No baseline yet -> establish one without regenerating (assume in sync).
    if old_state is None:
        state_mod.write_state(project_dir, new_state)
        return None

    # Fast, authoritative gate: identical per-slide hashes -> nothing to do.
    if old_state.get("slides") == new_state.get("slides"):
        if old_state != new_state:  # size drifted but content identical; refresh.
            state_mod.write_state(project_dir, new_state)
        return None

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
        summary = f"deck structure changed ({old_total} -> {total} slides); regenerated all {total}"
    else:
        target_slides = changed
        summary = f"regenerated slide(s) {', '.join(map(str, changed))} from your PowerPoint edit"

    # Suppress generate_slides' own stdout so the hook only injects the summary.
    with contextlib.redirect_stdout(io.StringIO()):
        generate_slides(out_pptx, project_dir, target_slides)

    state_mod.write_state(project_dir, new_state)
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Auto-sync deck code from PowerPoint edits")
    parser.add_argument(
        "--project-dir",
        default=".",
        help="Project directory, or a parent to scan (default: current dir)",
    )
    args = parser.parse_args()

    messages = []
    try:
        projects = _find_projects(Path(args.project_dir).resolve())
        for project in projects:
            try:
                summary = sync_project(project)
            except Exception as exc:  # never let one project block the prompt
                messages.append(f"[autosync] {project.name}: skipped ({type(exc).__name__}: {exc})")
                continue
            if summary:
                messages.append(f"[autosync] {project.name}: {summary}")
    except Exception as exc:
        messages.append(f"[autosync] skipped ({type(exc).__name__}: {exc})")

    if messages:
        print("\n".join(messages))
    # Always exit 0: a non-zero UserPromptSubmit hook can block the prompt.


if __name__ == "__main__":
    main()
