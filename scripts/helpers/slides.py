#!/usr/bin/env python3
"""Small helpers for reading slide XML metadata."""

import re
from pathlib import Path
from xml.etree import ElementTree as ET

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def sanitize_name(name: str) -> str:
    """Turn a slide title into a safe filename stem."""
    name = re.sub(r"[^\w\s-]", "", name)
    name = re.sub(r"[-\s]+", "_", name).strip("_").lower()
    name = re.sub(r"_+", "_", name)
    return name[:60]


def get_slide_title(slide_xml: Path) -> str:
    """Extract a title hint from the slide, ignoring footer/slide-number text."""
    tree = ET.parse(slide_xml)
    root = tree.getroot()

    # Collect candidate texts from each shape.
    candidates = []
    for sp in root.iter(f"{{{P}}}sp"):
        texts = [t.text.strip() for t in sp.iter(f"{{{A}}}t") if t.text and t.text.strip()]
        if not texts:
            continue
        primary = texts[0]
        # Skip footer and slide-number text.
        if primary == "PROTOPAPAS" or primary.isdigit():
            continue
        candidates.append(primary)

    return candidates[0] if candidates else "slide"


def get_slide_number(slide_xml: Path) -> int:
    """Extract the slide number from the slide-number placeholder shape."""
    tree = ET.parse(slide_xml)
    root = tree.getroot()
    for sp in root.iter(f"{{{P}}}sp"):
        xfrm = sp.find(f"{{{P}}}spPr/{{{A}}}xfrm")
        if xfrm is None:
            continue
        off = xfrm.find(f"{{{A}}}off")
        if off is None:
            continue
        try:
            x = int(off.get("x", 0)) / 914400
            y = int(off.get("y", 0)) / 914400
        except (ValueError, TypeError):
            continue
        # Match the slide-number placeholder position from design.py.
        if abs(x - 10.0) > 0.05 or abs(y - 7.0) > 0.05:
            continue
        texts = [t.text.strip() for t in sp.iter(f"{{{A}}}t") if t.text]
        if texts:
            try:
                return int(texts[0])
            except ValueError:
                return 0
    return 0
