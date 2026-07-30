#!/usr/bin/env python3
"""Leave a Claude-authored comment on a slide.

Comments are authored in the slide files as ``shapes.add_comment(...)`` calls, so
this script simply adds one to the target slide's file (inside the fenced comment
region, creating it if needed). It rides along on the next ``build_deck.py`` like
any other comment and behaves like a normal PowerPoint comment: a reviewer can
reply to it, resolve it, or delete it, and ``autosync.py`` mirrors that back.

Use it (per SKILL.md) when a slide edit is substantial, fixes a perceived error,
or addresses an existing reviewer comment -- not for routine formatting tweaks.

The new call is marked ``pending=True``: until the deck is rebuilt the comment
exists only in the code, and the deck -> code mirror would otherwise be unable to
tell it apart from a comment a human deleted in PowerPoint. The flag clears itself
on the first sync after a build.

Example::

    uv run python <skill>/scripts/add_comment.py \\
      --project-dir Session4_BERT --slide 71 \\
      --text "Corrected data size: BooksCorpus + Wikipedia is ~20-33 GB, not 40 TB (addresses reviewer note)."
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

_SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(_SCRIPTS))

from helpers.comments import (  # noqa: E402
    parse_comment_calls,
    patch_comment_region,
    render_comment_calls,
)

CLAUDE_AUTHOR_NAME = "Claude"


def _timestamp() -> str:
    now = datetime.now()
    return now.strftime("%Y-%m-%dT%H:%M:%S.") + f"{now.microsecond // 1000:03d}"


def _slide_file(project_dir: Path, slide: int):
    matches = sorted((project_dir / "slides").glob(f"s{slide:02d}_*.py"))
    return matches[0] if matches else None


def add_comment(project_dir: Path, slide: int, text: str, author: str) -> Path:
    project_dir = Path(project_dir)
    if not (project_dir / "slides").is_dir():
        raise SystemExit(
            f"error: {project_dir} does not look like a pptx-to-pypptx project "
            "(no slides/ directory)."
        )
    path = _slide_file(project_dir, slide)
    if path is None:
        raise SystemExit(f"error: no slide file for slide {slide} in {project_dir}/slides.")

    source = path.read_text(encoding="utf-8")
    threads = parse_comment_calls(source)
    threads.append(
        {
            "author": author,
            "created": _timestamp(),
            "text": text,
            "replies": [],
            "resolved": False,
            "pending": True,
        }
    )
    try:
        new_source = patch_comment_region(source, render_comment_calls(threads))
    except ValueError as exc:
        raise SystemExit(f"error: {exc}")
    if new_source is None:
        raise SystemExit("error: nothing to write (comment region unchanged).")
    path.write_text(new_source, encoding="utf-8")
    return path


def main() -> None:
    ap = argparse.ArgumentParser(description="Leave a Claude-authored comment on a slide.")
    ap.add_argument("--project-dir", required=True, help="Project directory (contains slides/).")
    ap.add_argument("--slide", required=True, type=int, help="1-based physical slide number.")
    ap.add_argument("--text", required=True, help="Comment text (keep it concise).")
    ap.add_argument(
        "--author",
        default=CLAUDE_AUTHOR_NAME,
        help=f"Comment author name (default: {CLAUDE_AUTHOR_NAME}).",
    )
    args = ap.parse_args()

    if not args.text.strip():
        raise SystemExit("error: --text is empty.")

    path = add_comment(Path(args.project_dir), args.slide, args.text, args.author)
    print(f"Added comment on slide {args.slide} -> {path}")
    print("It will be attached to out/<name>.pptx on the next build_deck.py run.")


if __name__ == "__main__":
    main()
