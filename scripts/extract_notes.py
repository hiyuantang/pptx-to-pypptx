#!/usr/bin/env python3
"""Extract speaker notes from a PPTX or generated slide files to Markdown.

Usage:
    uv run python extract_notes.py --target lecture.pptx \
        --output /tmp/speaker-notes.md
    uv run python extract_notes.py \
        --project-dir my-deck \
        --output my-deck/speaker_notes.md
"""

import argparse
import ast
import re
import tempfile
import zipfile
from pathlib import Path

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


def _render_markdown(title: str, slides: list[tuple[str, list[str]]]) -> str:
    lines = [f"# Speaker Notes: {title}", ""]

    for idx, (slide_title, notes) in enumerate(slides, start=1):
        lines.append(f"## Slide {idx}: {slide_title}")
        lines.append("")
        if notes:
            for note in notes:
                lines.append(note.strip())
                lines.append("")
        else:
            lines.append("_No speaker notes._")
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def generate_project_markdown(project_dir: Path) -> str:
    slide_dir = project_dir / "slides"
    if not slide_dir.exists():
        raise FileNotFoundError(f"Slides directory not found: {slide_dir}")

    slide_files = sorted(
        slide_dir.glob("s*.py"),
        key=_slide_sort_key,
    )
    slides = []
    for path in slide_files:
        text = path.read_text(encoding="utf-8")
        slides.append((extract_title(text) or path.stem, extract_notes(text)))
    return _render_markdown(project_dir.name, slides)


def generate_target_markdown(target: Path) -> str:
    if not target.exists():
        raise FileNotFoundError(f"Target PPTX not found: {target}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        with zipfile.ZipFile(target, "r") as zf:
            zf.extractall(tmp_path)

        slide_files = sorted((tmp_path / "ppt" / "slides").glob("slide*.xml"), key=_slide_sort_key)
        slides = []
        for path in slide_files:
            title = get_slide_title(path) or path.stem
            note = parse_slide_notes(path)
            slides.append((title, [note] if note else []))

    return _render_markdown(target.stem, slides)


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
    args = parser.parse_args()

    if args.target:
        target = Path(args.target)
        markdown = generate_target_markdown(target)
        default_output = target.with_name(f"{target.stem}-speaker-notes.md")
    else:
        project_dir = Path(args.project_dir)
        markdown = generate_project_markdown(project_dir)
        default_output = project_dir / "speaker_notes.md"

    output_path = Path(args.output) if args.output else default_output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(markdown, encoding="utf-8")
    print(f"Wrote speaker notes to {output_path}")


if __name__ == "__main__":
    main()
