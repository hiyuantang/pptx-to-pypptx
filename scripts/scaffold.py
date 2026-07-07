#!/usr/bin/env python3
"""Scaffold a new python-pptx project from a target PPTX.

Usage:
    uv run python scaffold.py --target "path/to/target.pptx" --output-dir my-deck

The generated deck is named after the output directory (build_deck.py writes
out/<output-dir-name>.pptx). The deck footer is auto-detected from the target
and written into lib/design.py as FOOTER_TEXT; generated slide chrome references
that constant, so editing it there updates the footer on every slide.
"""

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path

from helpers.assets import sync_assets
from helpers.pptx_utils import count_slides
from helpers.slide_codegen import detect_footer_text


def render_template(src: Path, dst: Path, replacements: dict) -> None:
    """Copy a template file, substituting placeholders."""
    text = src.read_text(encoding="utf-8")
    for key, value in replacements.items():
        text = text.replace(key, value)
    dst.write_text(text, encoding="utf-8")


def detect_footer(target: Path) -> str:
    """Infer the deck footer from the target's slide layouts (may be empty)."""
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(target, "r") as zf:
            zf.extractall(tmpdir)
        layouts = sorted((Path(tmpdir) / "ppt" / "slideLayouts").glob("slideLayout*.xml"))
        return detect_footer_text(layouts) or ""


def scaffold_project(target: Path, output_dir: Path) -> None:
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

    # Auto-detect the deck footer and bake it into design.py as FOOTER_TEXT.
    # design.py uses FOOTER_TEXT = __FOOTER_TEXT__ (no quotes), so we substitute
    # a repr'd value to stay safe for any footer string. Generated slide chrome
    # references d.FOOTER_TEXT, so editing it there updates every slide.
    replacements = {"__FOOTER_TEXT__": repr(detect_footer(target))}

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
    args = parser.parse_args()

    scaffold_project(Path(args.target), Path(args.output_dir))
