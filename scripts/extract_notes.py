#!/usr/bin/env python3
"""Extract speaker notes from generated slide files into a Markdown document.

Usage:
    uv run python extract_notes.py \
        --project-dir my-deck \
        --output my-deck/speaker_notes.md
"""

import argparse
import ast
import re
from pathlib import Path


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


def generate_markdown(project_dir: Path) -> str:
    slide_dir = project_dir / "slides"
    if not slide_dir.exists():
        raise FileNotFoundError(f"Slides directory not found: {slide_dir}")

    slide_files = sorted(
        slide_dir.glob("s*.py"),
        key=lambda p: int(p.stem.split("_", 1)[0][1:]) if p.stem[1:].split("_", 1)[0].isdigit() else p.stem,
    )

    lines = [f"# Speaker Notes: {project_dir.name}", ""]

    for idx, path in enumerate(slide_files, start=1):
        text = path.read_text(encoding="utf-8")
        title = extract_title(text) or path.stem
        notes = extract_notes(text)

        lines.append(f"## Slide {idx}: {title}")
        lines.append("")
        if notes:
            for note in notes:
                lines.append(note.strip())
                lines.append("")
        else:
            lines.append("_No speaker notes._")
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract speaker notes from slide files to Markdown.")
    parser.add_argument("--project-dir", required=True, help="Directory containing slides/")
    parser.add_argument(
        "--output",
        default=None,
        help="Output Markdown file (default: <project-dir>/speaker_notes.md)",
    )
    args = parser.parse_args()

    project_dir = Path(args.project_dir)
    output_path = Path(args.output) if args.output else project_dir / "speaker_notes.md"

    markdown = generate_markdown(project_dir)
    output_path.write_text(markdown, encoding="utf-8")
    print(f"Wrote speaker notes to {output_path}")


if __name__ == "__main__":
    main()
