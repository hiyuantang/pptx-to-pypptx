"""Round-trip sync state: track which version of the deck the code is in sync with.

Both ``build_deck.py`` (code -> deck) and the skill's ``scripts/autosync.py``
(deck -> code) call into this module, so their notion of "in sync" is computed
identically. The state is a small JSON marker written next to the project as
``.roundtrip_state.json``; it records the built/edited deck's size and a SHA-256
of every slide's XML.

``autosync.py``'s rule is simply: if ``out/<name>.pptx`` differs from the marker,
a human edited it in PowerPoint -> regenerate the changed slides. If it matches,
do nothing. ``build_deck.py`` re-stamps the marker after every build so its own
output is never mistaken for a human edit.

Only per-slide XML is hashed (small), not embedded media, so stamping and
detection stay cheap even on media-heavy decks. Note: PowerPoint may re-serialize
untouched slides on save, in which case detection over-reports and more slides
are regenerated than strictly necessary -- still correct, just heavier.
"""

import hashlib
import json
import re
import zipfile
from pathlib import Path

STATE_FILENAME = ".roundtrip_state.json"


def _slide_num(name: str) -> int:
    """``ppt/slides/slide12.xml`` -> ``12`` (matches generate_slides.py ordering)."""
    base = name.rsplit("/", 1)[-1]
    m = re.search(r"(\d+)", base)
    return int(m.group(1)) if m else 0


def compute_state(pptx: Path) -> dict:
    """Return ``{size, slide_count, slides: {slideN.xml: sha256}}`` for a .pptx."""
    pptx = Path(pptx)
    slides = {}
    with zipfile.ZipFile(pptx, "r") as zf:
        for name in zf.namelist():
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                base = name.rsplit("/", 1)[-1]
                slides[base] = hashlib.sha256(zf.read(name)).hexdigest()
    ordered = {k: slides[k] for k in sorted(slides, key=_slide_num)}
    return {
        "size": pptx.stat().st_size,
        "slide_count": len(ordered),
        "slides": ordered,
    }


def read_state(project_dir: Path) -> dict | None:
    """Load the marker, or ``None`` if it is missing or unreadable."""
    path = Path(project_dir) / STATE_FILENAME
    if not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None


def write_state(project_dir: Path, state: dict) -> None:
    """Persist the marker next to the project."""
    path = Path(project_dir) / STATE_FILENAME
    path.write_text(json.dumps(state, indent=2) + "\n", encoding="utf-8")


def stamp(project_dir: Path, pptx: Path) -> dict:
    """Compute the state for ``pptx`` and persist it as the marker; return it."""
    state = compute_state(pptx)
    write_state(project_dir, state)
    return state


def _ordered_hashes(state: dict) -> list[str]:
    slides = state.get("slides", {})
    return [slides[k] for k in sorted(slides, key=_slide_num)]


def changed_slides(old: dict | None, new: dict) -> list[int]:
    """Return the 1-based slide numbers whose XML differs between two states.

    Slide N is the N-th slide in ``slideN.xml`` numeric order -- the same
    convention ``generate_slides.py --slides N`` uses, so the returned numbers can
    be passed straight to it. If the slide counts differ (add/delete/reorder), all
    slides are returned so the caller can regenerate the whole deck.
    """
    new_hashes = _ordered_hashes(new)
    if not old:
        return []
    old_hashes = _ordered_hashes(old)
    if len(old_hashes) != len(new_hashes):
        return list(range(1, len(new_hashes) + 1))
    return [i for i, (o, n) in enumerate(zip(old_hashes, new_hashes), 1) if o != n]
