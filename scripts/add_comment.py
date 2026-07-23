#!/usr/bin/env python3
"""Leave a Claude-authored modern comment on a slide.

The generated deck preserves PowerPoint's modern threaded comments by storing
them under ``<project>/comments/`` and re-attaching them on every build (see
``template/lib/comments.py``). This script adds a *new* comment to that store so
Claude can annotate a slide with a concise note about a change it made -- the
comment then rides along on the next ``build_deck.py`` like any preserved one.

Use it (per SKILL.md) when a slide edit is substantial, fixes a perceived error,
or addresses an existing reviewer comment -- not for routine formatting tweaks.

The comment is pinned at slide level (no shape anchor); ``inject_comments`` in
``lib/comments.py`` rewrites its ``sldId`` to the rebuilt slide on each build.

Example::

    uv run python <skill>/scripts/add_comment.py \\
      --project-dir Session4_BERT --slide 71 \\
      --text "Corrected data size: BooksCorpus + Wikipedia is ~20-33 GB, not 40 TB (addresses reviewer note)."

Notes:
- The comment appears in ``out/<name>.pptx`` only after the next build.
- It behaves like a normal PowerPoint comment: a human reviewer can reply to it
  or delete it in PowerPoint, and ``autosync.py`` mirrors that change back into
  the store on the next deck task, so the edit sticks (the build will not
  resurrect a deleted comment).
"""

import argparse
import json
import uuid
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

# A fixed author identity so every Claude comment shares one authors.xml entry.
CLAUDE_AUTHOR_ID = "{0C1A0DE0-0000-4000-8000-000000000001}"
CLAUDE_AUTHOR_NAME = "Claude"
CLAUDE_AUTHOR_INITIALS = "AI"

# Records comments that are in the store but not yet in the built deck.
PENDING_FILE = ".pending.json"

# Namespaces used by modern (2018) PowerPoint comments.
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
_NS_PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"


def _guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def _cid() -> int:
    # Non-zero change-id for the thread; PowerPoint's real comments use large
    # unsigned ints here. Derive one from a fresh uuid so it varies per comment.
    return int(uuid.uuid4().hex[:8], 16) or 1


def _timestamp() -> str:
    now = datetime.now()
    return now.strftime("%Y-%m-%dT%H:%M:%S.") + f"{now.microsecond // 1000:03d}"


def _empty_authors_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p188:authorLst xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p188="{_NS_P188}">'
        "</p188:authorLst>"
    )


def _ensure_claude_author(authors_path: Path) -> None:
    """Add the Claude author entry to authors.xml if it is not already there."""
    if authors_path.exists():
        xml = authors_path.read_text(encoding="utf-8")
    else:
        xml = _empty_authors_xml()
    if f'id="{CLAUDE_AUTHOR_ID}"' in xml:
        return
    entry = (
        f'<p188:author id="{CLAUDE_AUTHOR_ID}" name="{CLAUDE_AUTHOR_NAME}" '
        f'initials="{CLAUDE_AUTHOR_INITIALS}" userId="{CLAUDE_AUTHOR_NAME}" '
        f'providerId="None"/>'
    )
    xml = xml.replace("</p188:authorLst>", entry + "</p188:authorLst>")
    authors_path.write_text(xml, encoding="utf-8")


def _record_pending(comments_dir: Path, cm_id: str) -> None:
    """Note a comment that exists in the store but not yet in the built deck.

    ``autosync.py`` mirrors the deck back into the store; without this record it
    could not tell a not-yet-built addition from a comment the human deleted in
    PowerPoint, and would silently destroy the former. Entries are dropped by
    autosync once the comment shows up in the deck (i.e. after a build).
    """
    path = comments_dir / PENDING_FILE
    try:
        data = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    except Exception:
        data = {}
    ids = [i for i in data.get("ids", []) if isinstance(i, str)]
    if cm_id not in ids:
        ids.append(cm_id)
    path.write_text(json.dumps({"ids": ids}, indent=2), encoding="utf-8")


def _cm_element(text: str, author_id: str, cm_id: str) -> str:
    # A single slide-level (unanchored) comment element. The markers live in
    # pc:sldMkLst, NOT ac:txMkLst: ac:txMkLst is the *text*-anchor list and the
    # schema requires it to carry spMk+txMk down to a text range; using it for a
    # slide-level pin makes PowerPoint offer to "repair" the file (the lenient
    # OOXML validator does not catch it). The a:/p188: prefixes resolve from the
    # enclosing cmLst; pc: is declared inline so this fragment can be appended
    # into an existing part. sldId is a placeholder; inject_comments() rewrites it.
    return (
        f'<p188:cm id="{cm_id}" authorId="{author_id}" created="{_timestamp()}">'
        f'<pc:sldMkLst xmlns:pc="{_NS_PC}"><pc:docMk/>'
        f'<pc:sldMk cId="{_cid()}" sldId="1"/></pc:sldMkLst>'
        "<p188:txBody><a:bodyPr/><a:lstStyle/>"
        f"<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{escape(text)}</a:t></a:r></a:p>"
        "</p188:txBody></p188:cm>"
    )


def _comment_file(cm_element: str) -> str:
    """Wrap a single cm element in a standalone comments part."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p188:cmLst xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p188="{_NS_P188}">'
        f"{cm_element}</p188:cmLst>"
    )


def add_comment(project_dir: Path, slide: int, text: str, author_id: str) -> Path:
    project_dir = Path(project_dir)
    slides_dir = project_dir / "slides"
    if not slides_dir.is_dir():
        raise SystemExit(
            f"error: {project_dir} does not look like a pptx-to-pypptx project "
            "(no slides/ directory)."
        )

    comments_dir = project_dir / "comments"
    comments_dir.mkdir(exist_ok=True)

    manifest_path = comments_dir / "manifest.json"
    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    else:
        manifest = {}
    manifest.setdefault("slides", {})
    manifest.setdefault("authors", "authors.xml")

    authors_path = comments_dir / manifest["authors"]
    _ensure_claude_author(authors_path)

    cm_id = _guid()
    cm = _cm_element(text, author_id, cm_id)
    existing = manifest["slides"].get(str(slide)) or []

    # PowerPoint allows exactly ONE comments part per slide (all threads share a
    # single cmLst). If the slide already has a part, append into it; creating a
    # second part on the same slide makes PowerPoint offer to "repair" the file.
    if existing:
        target = comments_dir / existing[0]
        xml = target.read_text(encoding="utf-8")
        xml = xml.replace("</p188:cmLst>", cm + "</p188:cmLst>", 1)
        target.write_text(xml, encoding="utf-8")
        manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        _record_pending(comments_dir, cm_id)
        return target

    fname = f"claudeComment_{uuid.uuid4().hex[:12]}.xml"
    (comments_dir / fname).write_text(_comment_file(cm), encoding="utf-8")
    manifest["slides"][str(slide)] = [fname]
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    _record_pending(comments_dir, cm_id)
    return comments_dir / fname


def main() -> None:
    ap = argparse.ArgumentParser(description="Leave a Claude-authored comment on a slide.")
    ap.add_argument("--project-dir", required=True, help="Project directory (contains slides/).")
    ap.add_argument("--slide", required=True, type=int, help="1-based physical slide number (slideN.xml order).")
    ap.add_argument("--text", required=True, help="Comment text (keep it concise).")
    ap.add_argument(
        "--author-id",
        default=CLAUDE_AUTHOR_ID,
        help="Author GUID (defaults to the shared Claude identity).",
    )
    args = ap.parse_args()

    if not args.text.strip():
        raise SystemExit("error: --text is empty.")

    path = add_comment(Path(args.project_dir), args.slide, args.text, args.author_id)
    print(f"Added comment on slide {args.slide} -> {path}")
    print("It will be attached to out/<name>.pptx on the next build_deck.py run.")


if __name__ == "__main__":
    main()
