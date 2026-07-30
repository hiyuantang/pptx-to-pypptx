"""Compile PowerPoint modern threaded comments from slide code into the built deck.

python-pptx cannot write comment parts at all, so comments are authored in the
slide files as ``shapes.add_comment(...)`` calls -- exactly like ``add_notes``,
which also writes to a part python-pptx owns separately. The helper only records
into a registry; this module turns that registry into real ``p188`` comment parts
and grafts them onto the saved package (``inject_comments``).

What the slide file carries is only what a human wrote: author, timestamp, text,
replies, resolved. Everything a comment needs but nobody wants to read lives
here:

* **Thread GUIDs are derived**, not stored -- ``uuid5(author_id + created)``. Same
  inputs produce the same GUID on every build, so PowerPoint keeps its per-thread
  read/unread state and the round-trip can tell an unbuilt comment apart from a
  deleted one, without a GUID in the slide file.
* **cId** (PowerPoint's change cookie) is derived from the GUID. It only has to be
  a stable non-zero uint; the value is not meaningful across decks.
* **Author identity** (userId / providerId / initials) is looked up in
  ``lib/authors.json`` by name, so the call site says ``author='Yuan Tang'``.

Every thread is pinned at slide level. Shape and text-range anchors from the
source deck are deliberately dropped: they reference source-deck shape
``creationId``s and character offsets that do not survive regeneration, so
carrying them would only preserve the *appearance* of an anchor.
"""

import json
import re
import uuid
import zipfile
from pathlib import Path

_AUTHORS_CT = "application/vnd.ms-powerpoint.authors+xml"
_COMMENTS_CT = "application/vnd.ms-powerpoint.comments+xml"
_AUTHORS_REL = "http://schemas.microsoft.com/office/2018/10/relationships/authors"
_COMMENTS_REL = "http://schemas.microsoft.com/office/2018/10/relationships/comments"

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
_NS_PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"

# Fixed namespace for derived ids: the same (author, timestamp) must always give
# the same GUID, across machines and runs, or PowerPoint sees new comments.
_ID_NS = uuid.UUID("6ba7b812-9dad-11d1-80b4-00c04fd430c8")

CLAUDE_AUTHOR = {
    "name": "Claude",
    "initials": "AI",
    "userId": "Claude",
    "providerId": "None",
}


def derive_id(author_id: str, created: str, salt: int = 0) -> str:
    """Stable thread GUID for an (author, timestamp) pair.

    ``salt`` disambiguates the vanishingly rare case of one author posting two
    comments in the same millisecond; ``compile_threads`` bumps it on collision.
    """
    seed = f"{author_id}|{created}" + (f"|{salt}" if salt else "")
    return "{" + str(uuid.uuid5(_ID_NS, seed)).upper() + "}"


def derive_author_id(name: str) -> str:
    """Stable author GUID for a name not present in ``lib/authors.json``."""
    return "{" + str(uuid.uuid5(_ID_NS, f"author|{name}")).upper() + "}"


def _derive_cid(guid: str) -> int:
    """Non-zero uint32 change cookie derived from a thread GUID."""
    digits = re.sub(r"[^0-9A-Fa-f]", "", guid)[:8]
    return int(digits, 16) or 1


def _esc(text: str) -> str:
    """Escape text for an XML text node (comment prose is arbitrary human input)."""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def load_authors(path) -> dict:
    """Load the deck's author table (``{guid: {name, initials, ...}}``)."""
    path = Path(path)
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    authors = data.get("authors") if isinstance(data, dict) else None
    if not isinstance(authors, dict):
        return {}
    return authors


def resolve_author(name: str, author_id, authors: dict) -> str:
    """Map an author name to its GUID, registering a derived one if unknown.

    Explicit ``author_id`` always wins -- the slide file needs it only when two
    people in the deck share a display name (which happens).
    """
    if author_id:
        authors.setdefault(author_id, {"name": name or "Unknown", "initials": "", "userId": "", "providerId": "None"})
        return author_id
    for guid, info in authors.items():
        if info.get("name") == name:
            return guid
    guid = derive_author_id(name)
    if name == CLAUDE_AUTHOR["name"]:
        authors[guid] = dict(CLAUDE_AUTHOR)
        return guid
    initials = "".join(part[0].upper() for part in str(name).split()[:2] if part)
    authors[guid] = {
        "name": name,
        "initials": initials,
        "userId": name,
        "providerId": "None",
    }
    return guid


def _txbody(text: str) -> str:
    return (
        "<p188:txBody><a:bodyPr/><a:lstStyle/>"
        f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{_esc(text)}</a:t></a:r></a:p>'
        "</p188:txBody>"
    )


def _reply_xml(reply: dict, parent_id: str, authors: dict) -> str:
    author_id = resolve_author(reply.get("author", ""), reply.get("author_id"), authors)
    created = reply.get("created", "")
    # Seeded with the parent id so two replies with identical author+timestamp on
    # different threads stay distinct.
    rid = derive_id(f"{parent_id}|{author_id}", created)
    return (
        f'<p188:reply id="{rid}" authorId="{author_id}" created="{created}">'
        f"{_txbody(reply.get('text', ''))}"
        "</p188:reply>"
    )


def compile_thread(thread: dict, sld_id, authors: dict, salt: int = 0) -> tuple:
    """Return ``(cm_id, cm_xml)`` for one thread pinned to ``sld_id``."""
    author_id = resolve_author(thread.get("author", ""), thread.get("author_id"), authors)
    created = thread.get("created", "")
    cm_id = derive_id(author_id, created, salt)
    cid = _derive_cid(cm_id)

    attrs = f'id="{cm_id}" authorId="{author_id}"'
    if thread.get("resolved"):
        attrs += ' status="resolved"'
    attrs += f' created="{created}"'
    if thread.get("resolved"):
        attrs += ' complete="100000"'

    # Element order is fixed by the schema: marker list, then replies, then the
    # comment's own body. Getting it wrong makes PowerPoint offer to repair the
    # file while the lenient OOXML validator stays quiet.
    marker = (
        f'<pc:sldMkLst xmlns:pc="{_NS_PC}"><pc:docMk/>'
        f'<pc:sldMk cId="{cid}" sldId="{sld_id}"/></pc:sldMkLst>'
    )
    replies = thread.get("replies") or []
    reply_xml = ""
    if replies:
        reply_xml = (
            "<p188:replyLst>"
            + "".join(_reply_xml(r, cm_id, authors) for r in replies)
            + "</p188:replyLst>"
        )
    body = _txbody(thread.get("text", ""))
    return cm_id, f"<p188:cm {attrs}>{marker}{reply_xml}{body}</p188:cm>"


def compile_threads(threads: list, sld_id, authors: dict) -> str:
    """Compile every thread for one slide into a single comments part.

    PowerPoint allows exactly ONE comments part per slide (all threads share a
    ``cmLst``); a second part on the same slide triggers a repair prompt.
    """
    seen = set()
    parts = []
    for thread in threads:
        salt = 0
        cm_id, xml = compile_thread(thread, sld_id, authors, salt)
        while cm_id in seen:  # same author, same millisecond
            salt += 1
            cm_id, xml = compile_thread(thread, sld_id, authors, salt)
        seen.add(cm_id)
        parts.append(xml)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p188:cmLst xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p188="{_NS_P188}">'
        + "".join(parts)
        + "</p188:cmLst>"
    )


def _authors_xml(authors: dict) -> str:
    entries = []
    for guid, info in authors.items():
        entries.append(
            f'<p188:author id="{guid}" name="{_quote(info.get("name", ""))}" '
            f'initials="{_quote(info.get("initials", ""))}" '
            f'userId="{_quote(info.get("userId", ""))}" '
            f'providerId="{_quote(info.get("providerId", "None"))}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p188:authorLst xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p188="{_NS_P188}">'
        + "".join(entries)
        + "</p188:authorLst>"
    )


def _quote(value: str) -> str:
    """Escape text for an XML attribute value."""
    return _esc(value).replace('"', "&quot;")


def _next_rid(rels_xml: str) -> str:
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return f"rId{max(ids, default=0) + 1}"


def _add_relationship(rels_xml: str, rid: str, rel_type: str, target: str) -> str:
    rel = f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/>'
    return rels_xml.replace("</Relationships>", rel + "</Relationships>")


def _add_override(ct_xml: str, part_name: str, content_type: str) -> str:
    if f'PartName="{part_name}"' in ct_xml:
        return ct_xml
    override = f'<Override PartName="{part_name}" ContentType="{content_type}"/>'
    return ct_xml.replace("</Types>", override + "</Types>")


def _slide_parts_by_sld_id(pres: str, pres_rels: str) -> dict:
    """``{sldId: 'ppt/slides/slideN.xml'}`` from the saved package."""
    rid_target = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", pres_rels):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        if attrs.get("Id") and attrs.get("Target"):
            rid_target[attrs["Id"]] = attrs["Target"]
    out = {}
    for tag in re.findall(r"<p:sldId\b[^>]*/>", pres):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        sldid, rid = attrs.get("id"), attrs.get("r:id")
        if not sldid or not rid:
            continue
        target = rid_target.get(rid, "")
        out[sldid] = target if target.startswith("ppt/") else "ppt/" + target.lstrip("/")
    return out


def inject_comments(pptx_path, registry: dict, authors_path=None) -> int:
    """Attach the comments recorded by ``shapes.add_comment`` to ``pptx_path``.

    ``registry`` maps a python-pptx ``slide.slide_id`` to its list of threads --
    the same id python-pptx writes as ``<p:sldId id=..>``, so no slide *position*
    is involved anywhere and inserting or reordering slides cannot mis-pin a
    comment. Returns the number of threads attached; a no-op for an empty
    registry.
    """
    if not registry:
        return 0
    authors = load_authors(authors_path) if authors_path else {}

    pptx_path = Path(pptx_path)
    with zipfile.ZipFile(pptx_path, "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    pres = items["ppt/presentation.xml"].decode("utf-8")
    pres_rels_name = "ppt/_rels/presentation.xml.rels"
    pres_rels = items[pres_rels_name].decode("utf-8")
    ct = items["[Content_Types].xml"].decode("utf-8")
    slide_parts = _slide_parts_by_sld_id(pres, pres_rels)

    attached = 0
    for slide_id, threads in registry.items():
        if not threads:
            continue
        slide_part = slide_parts.get(str(slide_id))
        if slide_part is None:
            continue
        part_xml = compile_threads(threads, slide_id, authors)
        fname = f"modernComment_{slide_id}.xml"
        part_name = f"ppt/comments/{fname}"
        items[part_name] = part_xml.encode("utf-8")
        ct = _add_override(ct, f"/{part_name}", _COMMENTS_CT)

        slide_fname = slide_part.split("/")[-1]
        slide_rels_name = f"ppt/slides/_rels/{slide_fname}.rels"
        raw = items.get(slide_rels_name)
        slide_rels_xml = raw.decode("utf-8") if raw is not None else (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>"
        )
        slide_rels_xml = _add_relationship(
            slide_rels_xml, _next_rid(slide_rels_xml), _COMMENTS_REL, f"../comments/{fname}"
        )
        items[slide_rels_name] = slide_rels_xml.encode("utf-8")
        attached += len(threads)

    if attached == 0:
        return 0

    # authors.xml is deck-wide and must list every id the comment parts reference
    # (resolve_author has by now added any that were missing from authors.json).
    items["ppt/authors.xml"] = _authors_xml(authors).encode("utf-8")
    ct = _add_override(ct, "/ppt/authors.xml", _AUTHORS_CT)
    if _AUTHORS_REL not in pres_rels:
        pres_rels = _add_relationship(
            pres_rels, _next_rid(pres_rels), _AUTHORS_REL, "authors.xml"
        )

    items[pres_rels_name] = pres_rels.encode("utf-8")
    items["[Content_Types].xml"] = ct.encode("utf-8")

    tmp_path = pptx_path.with_suffix(".pptx.comments.tmp")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    tmp_path.replace(pptx_path)
    return attached
