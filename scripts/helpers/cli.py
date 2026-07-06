#!/usr/bin/env python3
"""Shared CLI helpers for slide selection and PPTX inspection."""

import zipfile
from pathlib import Path


def parse_slide_arg(arg: str, total: int) -> list[int]:
    """Parse a slide argument like '14', '8-12', '4,5,9', or 'all'.

    Returns a sorted list of 1-based slide numbers clamped to [1, total].
    """
    if arg.lower() == "all":
        return list(range(1, total + 1))

    slides = set()
    for part in arg.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            slides.update(range(int(start), int(end) + 1))
        elif part:
            slides.add(int(part))
    return sorted(s for s in slides if 1 <= s <= total)


def count_slides_in_pptx(pptx_path: Path) -> int:
    """Count slide*.xml files inside a PPTX zip."""
    with zipfile.ZipFile(pptx_path, "r") as zf:
        return sum(
            1 for n in zf.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        )


