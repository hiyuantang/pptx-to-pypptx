#!/usr/bin/env python3
"""Detect a pptx-to-pypptx project directory and report its layout.

If --dir is given, check that directory only. Otherwise, check the current
working directory and one level of subdirectories.

Exit code:
    0 — a valid project was found
    1 — no valid project found
"""

import argparse
import json
import sys
from pathlib import Path


def _is_project_dir(path: Path) -> bool:
    return (
        path.is_dir()
        and (path / "build_deck.py").is_file()
        and (path / "slides").is_dir()
        and (path / "lib" / "shapes.py").is_file()
    )


def _find_project(start: Path) -> Path | None:
    if _is_project_dir(start):
        return start
    # Sort for deterministic results when several projects sit side by side.
    for child in sorted(start.iterdir()):
        if child.is_dir() and _is_project_dir(child):
            return child
    return None


def _output_pptx(project_dir: Path) -> tuple[str, Path | None]:
    """Return the expected output filename and the actual built file, if any.

    build_deck.py names the deck after its own directory — it writes
    ``out/<project-dir-name>.pptx`` (see template/build_deck.py). The file only
    exists after a build; if the directory or deck was renamed by hand, fall
    back to the newest ``out/*.pptx``.
    """
    expected = f"{project_dir.name}.pptx"
    out_dir = project_dir / "out"
    if (out_dir / expected).is_file():
        return expected, out_dir / expected
    built = sorted(out_dir.glob("*.pptx"), key=lambda p: p.stat().st_mtime)
    if built:
        return built[-1].name, built[-1]
    return expected, None


def main():
    parser = argparse.ArgumentParser(description="Detect a pptx-to-pypptx project")
    parser.add_argument(
        "--dir",
        type=Path,
        default=None,
        help="Directory to check (default: current dir and immediate subdirs)",
    )
    args = parser.parse_args()

    start = (args.dir or Path.cwd()).resolve()
    project_dir = _find_project(start)

    if project_dir is None:
        print(json.dumps({
            "found": False,
            "searched": str(start),
            "message": "No pptx-to-pypptx project found. Run scaffold.py to create one.",
        }, indent=2))
        sys.exit(1)

    slide_files = sorted((project_dir / "slides").glob("s*.py"))
    output_filename, out_pptx = _output_pptx(project_dir)
    backup_dir = project_dir / "backup"
    backup_files = sorted(backup_dir.glob("backup_*.pptx"), key=lambda p: p.name)

    result = {
        "found": True,
        "project_dir": str(project_dir),
        "output_filename": output_filename,
        "slide_count": len(slide_files),
        "slide_files": [f.name for f in slide_files],
        "has_backup_dir": backup_dir.is_dir(),
        "backup_count": len(backup_files),
        "latest_backup": str(backup_files[-1]) if backup_files else None,
        "output_pptx_exists": out_pptx is not None,
        "output_pptx_path": str(out_pptx) if out_pptx else None,
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
