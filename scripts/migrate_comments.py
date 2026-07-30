#!/usr/bin/env python3
"""One-time migration: move a project's comments from ``comments/`` into its slide files.

Older projects stored PowerPoint comments as verbatim ``p188`` XML parts under
``<project>/comments/``, keyed by slide *position* in ``manifest.json``. Comments
now live in the slide files as ``shapes.add_comment(...)`` calls, so they are
visible where the slide is edited and cannot drift onto the wrong slide when
slides are inserted or deleted.

The built deck (``out/<name>.pptx``) is the source of truth here, not the store:
``autosync.py`` mirrored the deck into the store on every task, so the deck holds
the current state of every thread including replies made in PowerPoint.

Dry run by default; pass ``--apply`` to write.

    uv run python <skill>/scripts/migrate_comments.py --project-dir session1 --apply
"""

import argparse
import shutil
import sys
from pathlib import Path

_SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(_SCRIPTS))

from detect_project import _output_pptx  # noqa: E402
from helpers.comments import (  # noqa: E402
    extract_authors,
    extract_comments,
    parse_comment_calls,
    patch_comment_region,
    render_comment_calls,
    write_authors_json,
)


def _slide_file(project_dir: Path, idx: int):
    matches = sorted((project_dir / "slides").glob(f"s{idx:02d}_*.py"))
    return matches[0] if matches else None


def migrate(project_dir: Path, apply: bool) -> int:
    project_dir = Path(project_dir).resolve()
    if not (project_dir / "slides").is_dir():
        raise SystemExit(f"error: {project_dir} has no slides/ directory.")

    _, out_pptx = _output_pptx(project_dir)
    if out_pptx is None or not out_pptx.exists():
        raise SystemExit(
            f"error: no built deck under {project_dir}/out. Build the deck first -- "
            "it is the source of truth for current comment state."
        )

    authors = extract_authors(out_pptx)
    comments = extract_comments(out_pptx)
    total = sum(len(v) for v in comments.values())
    print(f"{project_dir.name}: {total} thread(s) on {len(comments)} slide(s) in {out_pptx.name}")

    if apply:
        path = write_authors_json(authors, project_dir / "lib")
        if path:
            print(f"  wrote {path.relative_to(project_dir.parent)} ({len(authors)} authors)")

    written = skipped = 0
    for idx in sorted(comments):
        threads = comments[idx]
        path = _slide_file(project_dir, idx)
        if path is None:
            print(f"  slide {idx}: SKIP — no slide file")
            skipped += 1
            continue
        source = path.read_text(encoding="utf-8")
        if parse_comment_calls(source):
            print(f"  slide {idx}: already migrated ({path.name})")
            continue
        try:
            new_source = patch_comment_region(source, render_comment_calls(threads))
        except ValueError as exc:
            print(f"  slide {idx}: SKIP — {exc}")
            skipped += 1
            continue
        if new_source is None:
            continue
        print(f"  slide {idx}: {len(threads)} thread(s) -> {path.name}")
        if apply:
            path.write_text(new_source, encoding="utf-8")
        written += 1

    store = project_dir / "comments"
    if store.is_dir():
        print(f"  {'removing' if apply else 'would remove'} the old store {store.name}/")
        if apply:
            shutil.rmtree(store)

    # The comment digest is new in the round-trip marker, so drop the stale one;
    # the next autosync records a fresh baseline instead of comparing against
    # incompatible hashes.
    marker = project_dir / ".roundtrip_state.json"
    if marker.is_file():
        print(f"  {'clearing' if apply else 'would clear'} {marker.name} (new hash scheme)")
        if apply:
            marker.unlink()

    print(f"  {written} slide file(s) {'updated' if apply else 'to update'}, {skipped} skipped")
    if not apply:
        print("  (dry run — pass --apply to write)")
    return skipped


def main() -> None:
    ap = argparse.ArgumentParser(description="Migrate comments/ into slide files.")
    ap.add_argument("--project-dir", required=True, help="Project directory (contains slides/).")
    ap.add_argument("--apply", action="store_true", help="Write changes (default: dry run).")
    args = ap.parse_args()
    skipped = migrate(Path(args.project_dir), args.apply)
    sys.exit(1 if skipped else 0)


if __name__ == "__main__":
    main()
