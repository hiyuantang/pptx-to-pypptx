"""Shared .pptx utilities: slide counting and slide-range parsing.

Internal helper imported by the scripts in the parent directory; not run directly.
"""

import zipfile
from pathlib import Path


def count_slides(pptx: Path) -> int:
    """Count slide*.xml entries inside a .pptx zip."""
    with zipfile.ZipFile(pptx, "r") as zf:
        return sum(
            1 for n in zf.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        )


def parse_slide_range(arg: str, total: int, *, allow_all: bool = True) -> list[int]:
    """Parse a slide selection into a sorted list of 1-based slide numbers.

    Accepts a single number (``14``), a range (``8-12``), a comma list
    (``4,5,9``), or ``all`` (only when ``allow_all`` is set). Numbers are clamped
    to ``[1, total]``. Raises ``ValueError`` on malformed input, or on ``all``
    when ``allow_all`` is False.
    """
    if arg.strip().lower() == "all":
        if not allow_all:
            raise ValueError(
                "'all' is not supported here; specify slides like 4 | 2-5 | 3,7,9"
            )
        return list(range(1, total + 1))

    result = set()
    for part in arg.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            try:
                result.update(range(int(start), int(end) + 1))
            except ValueError:
                raise ValueError(f"Invalid slide range: {part!r}")
        else:
            try:
                result.add(int(part))
            except ValueError:
                raise ValueError(f"Invalid slide number: {part!r}")
    return sorted(n for n in result if 1 <= n <= total)
