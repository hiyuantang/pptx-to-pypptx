#!/usr/bin/env python3
"""Detect pptx-to-pypptx project directories and report their layout.

Searches --dir (or the current working directory) and recognizes a project at
the search dir itself or one level below. All matching projects are listed.

Exit code:
    0 — at least one project was found
    1 — no project found
"""

import argparse
import json
import sys
from pathlib import Path


def _is_project_dir(path: Path) -> bool:
    # Require the files a build actually needs: build_deck.py imports both
    # `from lib import design as d` and `from lib import shapes`, and lib must
    # be an importable package (lib/__init__.py).
    return (
        path.is_dir()
        and (path / "build_deck.py").is_file()
        and (path / "slides").is_dir()
        and (path / "lib" / "__init__.py").is_file()
        and (path / "lib" / "design.py").is_file()
        and (path / "lib" / "shapes.py").is_file()
    )


def _find_projects(start: Path) -> list[Path]:
    """Return every project at the start dir or one level below, sorted."""
    if not start.is_dir():
        return []
    found = []
    if _is_project_dir(start):
        found.append(start)
    # Sort for deterministic ordering when several projects sit side by side.
    for child in sorted(start.iterdir()):
        if _is_project_dir(child):
            found.append(child)
    return found


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


def _project_info(project_dir: Path) -> dict:
    slide_files = sorted((project_dir / "slides").glob("s*.py"))
    output_filename, out_pptx = _output_pptx(project_dir)
    backup_dir = project_dir / "backup"
    backup_files = sorted(backup_dir.glob("backup_*.pptx"), key=lambda p: p.name)
    return {
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


def main():
    parser = argparse.ArgumentParser(description="Detect pptx-to-pypptx projects")
    parser.add_argument(
        "--dir",
        type=Path,
        default=None,
        help="Directory to search (default: current dir and immediate subdirs)",
    )
    args = parser.parse_args()

    start = (args.dir or Path.cwd()).resolve()
    projects = _find_projects(start)

    result = {
        "found": bool(projects),
        "searched": str(start),
        "count": len(projects),
        "projects": [_project_info(p) for p in projects],
    }
    if not projects:
        result["message"] = "No pptx-to-pypptx project found. Run scaffold.py to create one."
    print(json.dumps(result, indent=2))
    sys.exit(0 if projects else 1)


if __name__ == "__main__":
    main()
