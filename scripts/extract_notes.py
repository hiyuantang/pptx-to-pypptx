#!/usr/bin/env python3
"""Extract speaker notes from a PPTX or generated slide files to Markdown.

Usage:
    uv run python extract_notes.py --target lecture.pptx \
        --output /tmp/speaker-notes.md
    uv run python extract_notes.py --target lecture.pptx \
        --output /tmp/lecture-source/speaker-notes.md \
        --slide-images-dir /tmp/lecture-source/slide-images
    uv run python extract_notes.py \
        --project-dir my-deck \
        --output my-deck/speaker_notes.md
"""

import argparse
import ast
import os
import re
import tempfile
import zipfile
from pathlib import Path

from helpers.pptx_utils import count_slides, parse_slide_range
from helpers.rendering import render_slides
from helpers.slide_meta import get_slide_title
from helpers.slide_xml import parse_slide_notes


def extract_title(text: str) -> str | None:
    """Extract the TITLE module-level constant."""
    m = re.search(r"^TITLE\s*=\s*(['\"])(.*?)\1", text, re.MULTILINE)
    if m:
        return ast.literal_eval(m.group(0).split("=", 1)[1].strip())
    return None


def extract_notes(text: str) -> list[str]:
    """Extract all string arguments from shapes.add_notes(slide, '...') calls."""
    notes = []
    pattern = re.compile(r"shapes\.add_notes\s*\(\s*slide\s*,\s*(['\"])(.*?)\s*(?<!\\)\1\s*\)", re.DOTALL)
    for m in pattern.finditer(text):
        raw = m.group(1) + m.group(2) + m.group(1)
        try:
            notes.append(ast.literal_eval(raw))
        except (SyntaxError, ValueError):
            notes.append(m.group(2))
    return notes


def _slide_sort_key(path: Path) -> int:
    match = re.search(r"(\d+)", path.stem)
    return int(match.group(1)) if match else 0


def _render_markdown(
    title: str,
    slides: list[tuple[int, str, list[str]]],
    preview_paths: dict[int, str] | None = None,
) -> str:
    lines = [f"# Speaker Notes: {title}", ""]

    for slide_num, slide_title, notes in slides:
        lines.append(f"<!-- lecture-source-slide: {slide_num} -->")
        lines.append(f"## Slide {slide_num}: {slide_title}")
        lines.append("")
        if preview_paths and slide_num in preview_paths:
            lines.append("<!-- lecture-source-preview:start -->")
            lines.append(
                f"![Temporary full-slide reference for slide {slide_num} — not a final "
                f"lecture-note asset]({preview_paths[slide_num]})"
            )
            lines.append("<!-- lecture-source-preview:end -->")
            lines.append("")
        if notes:
            for note in notes:
                lines.append(note.strip())
                lines.append("")
        else:
            lines.append("_No speaker notes._")
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def generate_project_markdown(
    project_dir: Path,
    selected_slides: list[int] | None = None,
) -> str:
    slide_dir = project_dir / "slides"
    if not slide_dir.exists():
        raise FileNotFoundError(f"Slides directory not found: {slide_dir}")

    slide_files = sorted(
        slide_dir.glob("s*.py"),
        key=_slide_sort_key,
    )
    selected = set(selected_slides or range(1, len(slide_files) + 1))
    slides = []
    for slide_num, path in enumerate(slide_files, start=1):
        if slide_num not in selected:
            continue
        text = path.read_text(encoding="utf-8")
        slides.append((slide_num, extract_title(text) or path.stem, extract_notes(text)))
    return _render_markdown(project_dir.name, slides)


def generate_target_markdown(
    target: Path,
    selected_slides: list[int] | None = None,
    preview_paths: dict[int, str] | None = None,
) -> str:
    if not target.exists():
        raise FileNotFoundError(f"Target PPTX not found: {target}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        with zipfile.ZipFile(target, "r") as zf:
            zf.extractall(tmp_path)

        slide_files = sorted((tmp_path / "ppt" / "slides").glob("slide*.xml"), key=_slide_sort_key)
        selected = set(selected_slides or range(1, len(slide_files) + 1))
        slides = []
        for slide_num, path in enumerate(slide_files, start=1):
            if slide_num not in selected:
                continue
            title = get_slide_title(path) or path.stem
            note = parse_slide_notes(path)
            slides.append((slide_num, title, [note] if note else []))

    return _render_markdown(target.stem, slides, preview_paths)


def _relative_preview_paths(
    rendered: dict[int, Path],
    markdown_path: Path,
) -> dict[int, str]:
    return {
        slide_num: Path(os.path.relpath(path, markdown_path.parent)).as_posix()
        for slide_num, path in rendered.items()
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract speaker notes from a PPTX or generated slide files to Markdown.")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--target", help="Source PPTX file (no generated project required)")
    source.add_argument("--project-dir", help="Directory containing generated slides/")
    parser.add_argument(
        "--output",
        default=None,
        help="Output Markdown file (defaults beside the selected source)",
    )
    parser.add_argument(
        "--slides",
        default="all",
        help="Slides to include: all | 4 | 2-5 | 3,7,9 (default: all)",
    )
    parser.add_argument(
        "--slide-images-dir",
        help=(
            "For --target, render temporary full-slide reference images here and link "
            "each one to its matching source section"
        ),
    )
    parser.add_argument(
        "--slide-image-dpi",
        type=int,
        default=120,
        help="DPI for temporary full-slide reference images (default: 120)",
    )
    args = parser.parse_args()

    if args.target:
        target = Path(args.target)
        default_output = target.with_name(f"{target.stem}-speaker-notes.md")
        total = count_slides(target)
    else:
        project_dir = Path(args.project_dir)
        default_output = project_dir / "speaker_notes.md"
        slide_dir = project_dir / "slides"
        total = len(list(slide_dir.glob("s*.py"))) if slide_dir.exists() else 0

    output_path = Path(args.output) if args.output else default_output
    if args.slide_image_dpi <= 0:
        parser.error("--slide-image-dpi must be positive")
    if args.slide_images_dir and not args.target:
        parser.error("--slide-images-dir requires --target")
    slides = parse_slide_range(args.slides, total)
    if not slides:
        parser.error("--slides did not select any slides")

    if args.target:
        preview_paths = None
        if args.slide_images_dir:
            rendered = render_slides(target, slides, Path(args.slide_images_dir), args.slide_image_dpi)
            preview_paths = _relative_preview_paths(rendered, output_path)
        markdown = generate_target_markdown(target, slides, preview_paths)
    else:
        markdown = generate_project_markdown(project_dir, slides)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(markdown, encoding="utf-8")
    print(f"Wrote speaker notes to {output_path}")
    if args.slide_images_dir:
        print(f"Wrote {len(slides)} temporary slide reference image(s) to {args.slide_images_dir}")


if __name__ == "__main__":
    main()
