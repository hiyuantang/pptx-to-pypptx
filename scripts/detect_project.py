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
    for child in start.iterdir():
        if child.is_dir() and _is_project_dir(child):
            return child
    return None


def _read_output_filename(project_dir: Path) -> str | None:
    build_deck = project_dir / "build_deck.py"
    try:
        text = build_deck.read_text(encoding="utf-8")
        for line in text.splitlines():
            if '__OUTPUT_FILENAME__' in line and '.pptx' in line:
                # Rendered template contains e.g. prs.save(out / "my-deck.pptx")
                for token in line.split('"'):
                    if token.endswith(".pptx"):
                        return token
                for token in line.split("'"):
                    if token.endswith(".pptx"):
                        return token
    except Exception:
        pass

    pyproject = project_dir / "pyproject.toml"
    try:
        text = pyproject.read_text(encoding="utf-8")
        for line in text.splitlines():
            if "name" in line and "=" in line:
                parts = line.split("=")
                if len(parts) == 2:
                    candidate = parts[1].strip().strip('"').strip("'")
                    if candidate:
                        return f"{candidate}.pptx"
    except Exception:
        pass
    return None


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
    out_pptx = project_dir / "out" / (_read_output_filename(project_dir) or "")
    backup_dir = project_dir / "backup"
    backup_files = sorted(backup_dir.glob("backup_*.pptx"), key=lambda p: p.name)

    result = {
        "found": True,
        "project_dir": str(project_dir),
        "output_filename": _read_output_filename(project_dir),
        "slide_count": len(slide_files),
        "slide_files": [f.name for f in slide_files],
        "has_backup_dir": backup_dir.is_dir(),
        "backup_count": len(backup_files),
        "latest_backup": str(backup_files[-1]) if backup_files else None,
        "output_pptx_exists": out_pptx.is_file(),
        "output_pptx_path": str(out_pptx) if out_pptx.is_file() else None,
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
