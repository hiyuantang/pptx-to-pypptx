#!/usr/bin/env python3
"""Recapture a project's base deck (lib/base.pptx) from an edited .pptx.

Run this after editing slide masters, layouts, or the theme in PowerPoint: it
refreshes lib/base.pptx so every subsequent build_deck.py uses the new
layouts/theme. Only lib/base.pptx is rewritten — your slides/ and lib/*.py are
left untouched (unlike scaffold.py, which regenerates the whole project).

Usage:
    uv run python recapture_base.py --project-dir my-deck
    uv run python recapture_base.py --project-dir my-deck --target path/to/edited.pptx

By default the source is the project's built deck (<project-dir>/out/<name>.pptx),
which is the deck you normally edit in PowerPoint.
"""

import argparse
from pathlib import Path

from helpers.pptx_utils import write_base_deck


def recapture_base(project_dir: Path, target: Path | None) -> None:
    lib_dir = project_dir / "lib"
    if not lib_dir.exists():
        raise FileNotFoundError(f"Not a scaffolded project (no lib/): {project_dir}")

    if target is None:
        target = project_dir / "out" / f"{project_dir.name}.pptx"
    if not target.exists():
        raise FileNotFoundError(
            f"Source deck not found: {target}\n"
            "Build the deck first, or pass --target <deck.pptx>."
        )

    dest = lib_dir / "base.pptx"
    write_base_deck(target, dest)
    print(f"Recaptured {dest}")
    print(f"  from: {target}")
    print("  build_deck.py will now use these masters/layouts/theme.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Refresh lib/base.pptx from an edited .pptx (non-destructive)."
    )
    parser.add_argument("--project-dir", required=True, help="Scaffolded project directory")
    parser.add_argument(
        "--target",
        default=None,
        help="Source .pptx to capture masters/layouts/theme from "
             "(default: <project-dir>/out/<name>.pptx).",
    )
    args = parser.parse_args()
    recapture_base(Path(args.project_dir), Path(args.target) if args.target else None)
