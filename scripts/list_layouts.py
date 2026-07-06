#!/usr/bin/env python3
"""List slide layouts in a .pptx with their indices.

Usage:
    uv run python list_layouts.py --target "path/to/target.pptx"
"""

import argparse
from pathlib import Path
from pptx import Presentation


def main() -> None:
    parser = argparse.ArgumentParser(description="List slide layouts in a PPTX.")
    parser.add_argument("--target", required=True, help="Target PPTX file")
    args = parser.parse_args()

    target = Path(args.target)
    if not target.exists():
        raise FileNotFoundError(f"Target PPTX not found: {target}")

    prs = Presentation(str(target))
    for i, layout in enumerate(prs.slide_layouts):
        print(f"{i}: {layout.name}")


if __name__ == "__main__":
    main()
