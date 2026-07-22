"""Re-attach PowerPoint modern threaded comments to the built deck.

python-pptx drops comment parts, so scaffold.py copies them into
``<project>/comments/`` (``authors.xml``, the ``modernComment_*.xml`` parts, and
``manifest.json`` mapping the 1-based deck slide index -> comment filenames).
``inject_comments()`` copies those parts back into the freshly built package,
registers their content types, wires the presentation->authors and
slide->comment relationships, and rewrites each comment's ``<pc:sldMk sldId=..>``
marker to the rebuilt slide's id so the thread anchors to the right slide.

Preserved: comment text, author, timestamp, thread replies, slide association.
Not preserved: shape-level anchoring (``ac:spMk``), since rebuilt shape ids
differ; those comments fall back to a slide-level pin (the on-slide position is
still carried by ``p188:pos``).

The whole step is a no-op when the project has no ``comments/manifest.json``.
"""

import json
import re
import zipfile
from pathlib import Path

_AUTHORS_CT = "application/vnd.ms-powerpoint.authors+xml"
_COMMENTS_CT = "application/vnd.ms-powerpoint.comments+xml"
_AUTHORS_REL = "http://schemas.microsoft.com/office/2018/10/relationships/authors"
_COMMENTS_REL = "http://schemas.microsoft.com/office/2018/10/relationships/comments"


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


def inject_comments(pptx_path, comments_dir) -> int:
    """Re-attach preserved comments to ``pptx_path`` in place. Returns count."""
    comments_dir = Path(comments_dir)
    manifest_path = comments_dir / "manifest.json"
    if not manifest_path.exists():
        return 0
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    slides_map = manifest.get("slides") or {}
    if not slides_map:
        return 0

    pptx_path = Path(pptx_path)
    with zipfile.ZipFile(pptx_path, "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    pres = items["ppt/presentation.xml"].decode("utf-8")
    pres_rels_name = "ppt/_rels/presentation.xml.rels"
    pres_rels = items[pres_rels_name].decode("utf-8")
    ct = items["[Content_Types].xml"].decode("utf-8")

    # Deck position (1-based, document order of <p:sldId>) -> (sldId, slide part).
    rid_target = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", pres_rels):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        if attrs.get("Id") and attrs.get("Target"):
            rid_target[attrs["Id"]] = attrs["Target"]
    position_to_slide = {}
    idx = 0
    for tag in re.findall(r"<p:sldId\b[^>]*/>", pres):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        sldid, rid = attrs.get("id"), attrs.get("r:id")
        if not sldid or not rid:
            continue
        idx += 1
        target = rid_target.get(rid, "")
        slide_part = target if target.startswith("ppt/") else "ppt/" + target.lstrip("/")
        position_to_slide[idx] = (sldid, slide_part)

    injected = 0

    authors_file = manifest.get("authors")
    if authors_file and (comments_dir / authors_file).exists():
        items["ppt/authors.xml"] = (comments_dir / authors_file).read_bytes()
        ct = _add_override(ct, "/ppt/authors.xml", _AUTHORS_CT)
        if _AUTHORS_REL not in pres_rels:
            pres_rels = _add_relationship(
                pres_rels, _next_rid(pres_rels), _AUTHORS_REL, "authors.xml"
            )

    for pos_str, files in slides_map.items():
        pos = int(pos_str)
        if pos not in position_to_slide:
            continue
        sldid, slide_part = position_to_slide[pos]
        slide_fname = slide_part.split("/")[-1]
        slide_rels_name = f"ppt/slides/_rels/{slide_fname}.rels"
        raw = items.get(slide_rels_name)
        slide_rels_xml = raw.decode("utf-8") if raw is not None else (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>"
        )
        for fname in files:
            src = comments_dir / fname
            if not src.exists():
                continue
            data = src.read_text(encoding="utf-8")
            # Anchor the thread to the rebuilt slide (python-pptx assigns fresh
            # slide ids, so the captured sldMk marker would otherwise dangle).
            data = re.sub(r'(<pc:sldMk\b[^>]*\bsldId=")\d+(")', rf"\g<1>{sldid}\g<2>", data)
            part_name = f"ppt/comments/{fname}"
            items[part_name] = data.encode("utf-8")
            ct = _add_override(ct, f"/{part_name}", _COMMENTS_CT)
            rid = _next_rid(slide_rels_xml)
            slide_rels_xml = _add_relationship(
                slide_rels_xml, rid, _COMMENTS_REL, f"../comments/{fname}"
            )
            injected += 1
        items[slide_rels_name] = slide_rels_xml.encode("utf-8")

    if injected == 0:
        return 0

    items[pres_rels_name] = pres_rels.encode("utf-8")
    items["[Content_Types].xml"] = ct.encode("utf-8")

    tmp_path = pptx_path.with_suffix(".pptx.comments.tmp")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    tmp_path.replace(pptx_path)
    return injected
