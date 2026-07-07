"""Slide metadata helpers: derive a title hint and a safe filename stem.

Internal helper imported by the scripts in the parent directory; not run directly.
"""

import re
from pathlib import Path
from xml.etree import ElementTree as ET

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# PowerPoint's slide-number field renders as this marker (‹#›) when its value
# isn't cached as a plain integer; it is never a slide title.
SLDNUM_MARKER = "‹#›"


def sanitize_name(name: str) -> str:
    """Turn a slide title into a safe filename stem."""
    name = re.sub(r"[^\w\s-]", "", name)
    name = re.sub(r"[-\s]+", "_", name).strip("_").lower()
    name = re.sub(r"_+", "_", name)
    return name[:60]


def get_slide_title(slide_xml: Path) -> str:
    """Extract a title hint from the slide, skipping pure slide-number text."""
    tree = ET.parse(slide_xml)
    root = tree.getroot()

    # Return the first shape's first line of text, skipping slide-number shapes:
    # both the sldNum placeholder and any leftover ‹#› / bare-digit field text.
    for sp in root.iter(f"{{{P}}}sp"):
        ph = sp.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph")
        if ph is not None and ph.get("type") == "sldNum":
            continue
        texts = [t.text.strip() for t in sp.iter(f"{{{A}}}t") if t.text and t.text.strip()]
        if not texts:
            continue
        primary = texts[0]
        if primary.isdigit() or primary == SLDNUM_MARKER:
            continue
        return primary

    return "slide"
