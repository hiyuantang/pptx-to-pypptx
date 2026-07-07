#!/usr/bin/env python3
"""Scaffold a new python-pptx project from a target PPTX.

Usage:
    uv run python scaffold.py \
        --target "path/to/target.pptx" \
        --output-dir my-deck \
        [--project-name my-deck] \
        [--output-filename my-deck.pptx] \
        [--footer-text "My Name"]
"""

import argparse
import shutil
from pathlib import Path

from helpers.assets import sync_assets
from helpers.pptx_utils import count_slides


def render_template(src: Path, dst: Path, replacements: dict) -> None:
    """Copy a template file, substituting placeholders."""
    text = src.read_text(encoding="utf-8")
    for key, value in replacements.items():
        text = text.replace(key, value)
    dst.write_text(text, encoding="utf-8")


def scaffold_project(
    target: Path,
    output_dir: Path,
    project_name: str,
    output_filename: str,
    footer_text: str,
) -> None:
    if not target.exists():
        raise FileNotFoundError(f"Target PPTX not found: {target}")

    output_dir.mkdir(parents=True, exist_ok=True)
    slides_dir = output_dir / "slides"
    lib_dir = output_dir / "lib"
    out_dir = output_dir / "out"
    slides_dir.mkdir(exist_ok=True)
    lib_dir.mkdir(exist_ok=True)
    out_dir.mkdir(exist_ok=True)

    # Locate template directory relative to this script
    script_dir = Path(__file__).resolve().parent
    template_dir = script_dir.parent / "template"
    if not template_dir.exists():
        raise FileNotFoundError(f"Template directory not found: {template_dir}")

# Clear previous scaffold output so re-runs stay consistent.
    # Never delete .pptx files that a human may have placed in the output dir.
    for f in slides_dir.glob("s*.py"):
        f.unlink()
    for f in lib_dir.glob("*.py"):
        f.unlink()
    for f in output_dir.glob("*.py"):
        f.unlink()

    # Sync media assets from the target PPTX, deduplicating by content hash.
    if (output_dir / "assets").exists():
        shutil.rmtree(output_dir / "assets")
    sync_assets(target, output_dir)

    # Create a backup directory where successful builds will be archived.
    # This lets the user roll back up to 10 previous generated decks.
    backup_dir = output_dir / "backup"
    backup_dir.mkdir(exist_ok=True)

    replacements = {
        "__PROJECT_NAME__": project_name,
        "__OUTPUT_FILENAME__": output_filename,
        "__FOOTER_TEXT__": footer_text,
    }

    # Copy top-level templates
    render_template(template_dir / "build_deck.py", output_dir / "build_deck.py", replacements)

    # Copy lib templates
    render_template(template_dir / "lib" / "__init__.py", lib_dir / "__init__.py", replacements)
    render_template(template_dir / "lib" / "design.py", lib_dir / "design.py", replacements)
    render_template(template_dir / "lib" / "shapes.py", lib_dir / "shapes.py", replacements)

    # Make generated scripts executable
    (output_dir / "build_deck.py").chmod(0o755)

    slide_count = count_slides(target)

    print(f"Scaffolded python-pptx project at {output_dir}")
    print(f"  slides directory: {slides_dir} (empty; run generate_slides.py to populate)")
    print(f"  target slides: {slide_count}")
    print(f"  assets: {output_dir / 'assets'}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scaffold a python-pptx project from a target PPTX")
    parser.add_argument("--target", required=True, help="Target PPTX file")
    parser.add_argument("--output-dir", required=True, help="Output directory for the new project")
    parser.add_argument("--project-name", default="my-deck", help="Project name (default: my-deck)")
    parser.add_argument(
        "--output-filename",
        default=None,
        help="Generated PPTX filename (default: <project-name>.pptx)",
    )
    parser.add_argument("--footer-text", default="", help="Footer text (default: empty)")
    args = parser.parse_args()

    output_filename = args.output_filename or f"{args.project_name}.pptx"
    scaffold_project(
        Path(args.target),
        Path(args.output_dir),
        args.project_name,
        output_filename,
        args.footer_text,
    )
