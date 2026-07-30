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
detection stay cheap even on media-heavy decks.

The XML is canonicalized (C14N) before hashing. python-pptx (which ``build_deck``
uses to write the deck) and PowerPoint (which the human edits with) serialize the
*same* slide content with different formatting -- quote style, the newline after
the XML declaration, attribute/namespace ordering. Hashing the raw bytes would
report every slide as changed the first time a human saves in PowerPoint, even
untouched ones. Canonicalizing collapses those cosmetic differences so only
genuinely edited slides are detected.
"""

import hashlib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

STATE_FILENAME = ".roundtrip_state.json"

# Bump when the hashing scheme changes so a stale baseline is recomputed rather
# than silently compared against incompatible hashes.
HASH_SCHEME = "c14n-sha256+cm1"

# Comments live in their own part, so a reply or deletion leaves every slide hash
# untouched. They are digested separately (see _comment_digest) and autosync uses
# changed_comment_slides() to patch just the comment region of those slide files.
_COMMENTS_REL = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
_P188 = "{http://schemas.microsoft.com/office/powerpoint/2018/8/main}"
_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _slide_digest(raw: bytes) -> str:
    """SHA-256 of a slide's canonicalized XML.

    Canonicalization (C14N) normalizes serializer-specific formatting so
    python-pptx's and PowerPoint's serializations of identical content hash the
    same. Falls back to the raw bytes if the XML can't be parsed, so a single odd
    slide over-reports rather than crashing the whole sync.
    """
    try:
        canonical = ET.canonicalize(raw.decode("utf-8"))
        return hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    except Exception:
        return hashlib.sha256(raw).hexdigest()


def _slide_num(name: str) -> int:
    """``ppt/slides/slide12.xml`` -> ``12`` (matches generate_slides.py ordering)."""
    base = name.rsplit("/", 1)[-1]
    m = re.search(r"(\d+)", base)
    return int(m.group(1)) if m else 0


def _text_of(txbody) -> str:
    """Concatenated text of a ``p188:txBody`` (``None`` -> empty)."""
    if txbody is None:
        return ""
    return "".join(t.text or "" for t in txbody.iter(f"{_A_NS}t"))


def _comment_digest(raw: bytes) -> str:
    """SHA-256 over a comment part's *meaning*, ignoring machine-facing fields.

    Only what a human wrote is hashed -- author, timestamp, resolved state, text,
    and replies. Thread GUIDs, ``cId`` and ``sldId`` are excluded: they are
    derived or per-build, so hashing them would report a change on every rebuild
    and trigger pointless rewrites. Threads are sorted so a cosmetic reorder in
    PowerPoint is not mistaken for an edit.
    """
    try:
        root = ET.fromstring(raw.decode("utf-8"))
    except Exception:
        return hashlib.sha256(raw).hexdigest()
    entries = []
    for cm in root.findall(f"{_P188}cm"):
        replies = [
            (
                reply.get("authorId", ""),
                reply.get("created", ""),
                _text_of(reply.find(f"{_P188}txBody")),
            )
            for reply in cm.findall(f"{_P188}replyLst/{_P188}reply")
        ]
        entries.append((
            cm.get("authorId", ""),
            cm.get("created", ""),
            cm.get("status", "") or "",
            _text_of(cm.find(f"{_P188}txBody")),
            tuple(replies),
        ))
    entries.sort()
    return hashlib.sha256(repr(entries).encode("utf-8")).hexdigest()


def _comment_part_for_slide(zf, slide_name: str):
    """Resolve a slide's comment part name via its rels, or ``None``."""
    base = slide_name.rsplit("/", 1)[-1]
    rels_name = f"ppt/slides/_rels/{base}.rels"
    try:
        rels = zf.read(rels_name).decode("utf-8")
    except KeyError:
        return None
    for tag in re.findall(r"<Relationship\b[^>]*/>", rels):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', tag))
        if attrs.get("Type") == _COMMENTS_REL and attrs.get("Target"):
            target = attrs["Target"].split("/")[-1]
            return f"ppt/comments/{target}"
    return None


def compute_state(pptx: Path) -> dict:
    """Return ``{hash, size, slide_count, slides, comments}`` for a .pptx.

    ``slides`` maps ``slideN.xml`` -> canonicalized-XML digest; ``comments`` maps
    the same keys -> comment-thread digest, present only for slides that have
    comments.
    """
    pptx = Path(pptx)
    slides = {}
    comments = {}
    with zipfile.ZipFile(pptx, "r") as zf:
        for name in zf.namelist():
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                base = name.rsplit("/", 1)[-1]
                slides[base] = _slide_digest(zf.read(name))
                part = _comment_part_for_slide(zf, name)
                if part:
                    try:
                        comments[base] = _comment_digest(zf.read(part))
                    except KeyError:
                        pass
    ordered = {k: slides[k] for k in sorted(slides, key=_slide_num)}
    ordered_comments = {k: comments[k] for k in sorted(comments, key=_slide_num)}
    return {
        "hash": HASH_SCHEME,
        "size": pptx.stat().st_size,
        "slide_count": len(ordered),
        "slides": ordered,
        "comments": ordered_comments,
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


def _ordered_comment_hashes(state: dict) -> list[str]:
    """Comment digests in slide order, ``""`` for slides with no comments."""
    slides = state.get("slides", {})
    comments = state.get("comments", {})
    return [comments.get(k, "") for k in sorted(slides, key=_slide_num)]


def changed_comment_slides(old: dict | None, new: dict) -> list[int]:
    """Return the 1-based slide numbers whose *comments* differ between states.

    Comments live in a separate part, so a reply or deletion in PowerPoint leaves
    the slide digests identical and ``changed_slides`` returns nothing. autosync
    uses this to patch only the comment region of the affected slide files,
    leaving hand-written slide code untouched.
    """
    if not old:
        return []
    new_hashes = _ordered_comment_hashes(new)
    old_hashes = _ordered_comment_hashes(old)
    if len(old_hashes) != len(new_hashes):
        return list(range(1, len(new_hashes) + 1))
    return [i for i, (o, n) in enumerate(zip(old_hashes, new_hashes), 1) if o != n]
