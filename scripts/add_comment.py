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

# Namespaces used by modern (2018) PowerPoint comments.
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
_NS_AC = "http://schemas.microsoft.com/office/drawing/2013/main/command"
_NS_PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"

# Where the comment pin sits on the slide (EMU). Top-right, clear of most body
# content, so Claude's notes cluster in a consistent spot.
_POS_X = 11000000
_POS_Y = 400000


def _guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


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


def _comment_xml(text: str, author_id: str) -> str:
    # sldId is a placeholder; inject_comments() rewrites it to the rebuilt slide.
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p188:cmLst xmlns:a="{_NS_A}" xmlns:r="{_NS_R}" xmlns:p188="{_NS_P188}">'
        f'<p188:cm id="{_guid()}" authorId="{author_id}" created="{_timestamp()}">'
        f'<ac:txMkLst xmlns:ac="{_NS_AC}">'
        f'<pc:docMk xmlns:pc="{_NS_PC}"/>'
        f'<pc:sldMk xmlns:pc="{_NS_PC}" cId="0" sldId="1"/>'
        "</ac:txMkLst>"
        f'<p188:pos x="{_POS_X}" y="{_POS_Y}"/>'
        "<p188:txBody><a:bodyPr/><a:lstStyle/>"
        f"<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{escape(text)}</a:t></a:r></a:p>"
        "</p188:txBody></p188:cm></p188:cmLst>"
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

    fname = f"claudeComment_{uuid.uuid4().hex[:12]}.xml"
    (comments_dir / fname).write_text(_comment_xml(text, author_id), encoding="utf-8")

    manifest["slides"].setdefault(str(slide), []).append(fname)
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

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
