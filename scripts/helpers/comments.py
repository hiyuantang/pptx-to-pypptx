"""Read PowerPoint comments out of a deck and render them as slide-file code.

Comments are authored in the slide files as ``shapes.add_comment(...)`` calls, so
this module is the deck -> code direction: parse the ``p188`` comment parts into
plain thread dicts, render them as a delimited region of Python, and splice that
region into a slide file without disturbing anything else in it.

The region is fenced by ``REGION_START`` / ``REGION_END`` so ``autosync.py`` can
replace it with a bounded find-and-replace. That matters because comments change
on a different clock than slides: a reviewer replying in PowerPoint leaves the
slide's own XML untouched, so the file must be patched, not regenerated -- any
hand-written slide code around the fence has to survive.

Round-trip note: shape and text-range anchors are read but discarded. They point
at source-deck ``creationId``s and character offsets that do not survive
regeneration, so every thread becomes a slide-level pin (which is what those
anchors already degraded to).
"""

import ast
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Slide -> comment-thread relationship used by modern (2018) PowerPoint comments.
COMMENTS_REL_TYPE = "http://schemas.microsoft.com/office/2018/10/relationships/comments"

_P188 = "{http://schemas.microsoft.com/office/powerpoint/2018/8/main}"
_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Fence around the generated calls. autosync replaces everything between these
# two lines and nothing else, so hand-written slide code is never at risk.
REGION_START = "# --- comments (managed; edit in PowerPoint, synced automatically) ---"
REGION_END = "# --- end comments ---"


def _slide_key(name: str) -> int:
    m = re.search(r"(\d+)", name)
    return int(m.group(1)) if m else 0


def _text_of(txbody) -> str:
    if txbody is None:
        return ""
    return "".join(t.text or "" for t in txbody.iter(f"{_A_NS}t"))


def extract_authors(target: Path) -> dict:
    """``{author_guid: {name, initials, userId, providerId}}`` from ppt/authors.xml."""
    authors: dict = {}
    with zipfile.ZipFile(Path(target), "r") as zf:
        if "ppt/authors.xml" not in zf.namelist():
            return authors
        try:
            root = ET.fromstring(zf.read("ppt/authors.xml").decode("utf-8"))
        except Exception:
            return authors
        for author in root.findall(f"{_P188}author"):
            guid = author.get("id")
            if not guid:
                continue
            authors[guid] = {
                "name": author.get("name", ""),
                "initials": author.get("initials", ""),
                "userId": author.get("userId", ""),
                "providerId": author.get("providerId", "None"),
            }
    return authors


def _parse_part(raw: bytes, authors: dict, ambiguous: set) -> list:
    """Parse one comment part into ``[thread]`` in document order."""
    try:
        root = ET.fromstring(raw.decode("utf-8"))
    except Exception:
        return []
    threads = []
    for cm in root.findall(f"{_P188}cm"):
        author_id = cm.get("authorId", "")
        name = authors.get(author_id, {}).get("name", "") or "Unknown"
        thread = {
            "author": name,
            "created": cm.get("created", ""),
            "text": _text_of(cm.find(f"{_P188}txBody")),
            "resolved": cm.get("status", "") == "resolved",
            "replies": [],
        }
        # Only pin down the GUID when the display name is genuinely shared, so the
        # common case stays a readable author= with no machine-facing noise.
        if name in ambiguous:
            thread["author_id"] = author_id
        for reply in cm.findall(f"{_P188}replyLst/{_P188}reply"):
            r_author_id = reply.get("authorId", "")
            r_name = authors.get(r_author_id, {}).get("name", "") or "Unknown"
            entry = {
                "author": r_name,
                "created": reply.get("created", ""),
                "text": _text_of(reply.find(f"{_P188}txBody")),
            }
            if r_name in ambiguous:
                entry["author_id"] = r_author_id
            thread["replies"].append(entry)
        threads.append(thread)
    return threads


def extract_comments(target: Path) -> dict:
    """``{slide_index: [thread]}`` for a deck, 1-based in ``slideN.xml`` order.

    Slides with no comments are absent; a deck with none returns ``{}``.
    """
    target = Path(target)
    authors = extract_authors(target)
    names = [info.get("name", "") for info in authors.values()]
    ambiguous = {n for n in names if n and names.count(n) > 1}

    out: dict = {}
    with zipfile.ZipFile(target, "r") as zf:
        in_zip = set(zf.namelist())
        if not any(re.match(r"ppt/comments/.*\.xml$", n) for n in in_zip):
            return out
        slide_files = sorted(
            (n for n in in_zip if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=_slide_key,
        )
        for idx, slide_name in enumerate(slide_files, start=1):
            rels_name = f"ppt/slides/_rels/{slide_name.split('/')[-1]}.rels"
            if rels_name not in in_zip:
                continue
            rels_xml = zf.read(rels_name).decode("utf-8")
            threads = []
            for tag in re.findall(r"<Relationship\b[^>]*/>", rels_xml):
                attrs = dict(re.findall(r'(\w+)="([^"]*)"', tag))
                if attrs.get("Type") != COMMENTS_REL_TYPE or not attrs.get("Target"):
                    continue
                part = f"ppt/comments/{attrs['Target'].split('/')[-1]}"
                if part in in_zip:
                    threads.extend(_parse_part(zf.read(part), authors, ambiguous))
            if threads:
                out[idx] = threads
    return out


def write_authors_json(authors: dict, lib_dir: Path) -> Path | None:
    """Persist the deck-wide author table next to the other ``lib/`` data.

    Deliberately not in the slide files: ``userId``/``providerId``/initials are
    machine-facing, and two people in one deck can share a display name, so this
    table is what makes ``author='Yuan Tang'`` resolvable at build time.
    """
    lib_dir = Path(lib_dir)
    path = lib_dir / "authors.json"
    if not authors:
        if path.is_file():
            path.unlink()
        return None
    lib_dir.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"authors": authors}, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return path


def merge_authors_json(authors: dict, lib_dir: Path) -> Path | None:
    """Refresh ``lib/authors.json`` from the deck, keeping ids it no longer lists.

    A deck whose last comment by someone was deleted drops them from
    ``authors.xml``, but an unbuilt ``pending`` comment in the code may still name
    them, so existing entries are preserved rather than replaced wholesale.
    """
    lib_dir = Path(lib_dir)
    path = lib_dir / "authors.json"
    existing: dict = {}
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data.get("authors"), dict):
                existing = data["authors"]
        except (json.JSONDecodeError, OSError):
            existing = {}
    merged = {**existing, **authors}
    return write_authors_json(merged, lib_dir)


def _reply_literal(reply: dict) -> str:
    parts = [
        f"'author': {reply.get('author', '')!r}",
        f"'created': {reply.get('created', '')!r}",
        f"'text': {reply.get('text', '')!r}",
    ]
    if reply.get("author_id"):
        parts.append(f"'author_id': {reply['author_id']!r}")
    return "{" + ", ".join(parts) + "}"


def render_comment_calls(threads: list, indent: str = "    ") -> str:
    """Render threads as the fenced ``shapes.add_comment(...)`` region.

    Every string goes through ``repr()``, which is valid Python for any content --
    quotes, curly apostrophes, newlines, emoji -- so reviewer prose cannot break
    the file it lands in.
    """
    if not threads:
        return ""
    lines = [f"{indent}{REGION_START}"]
    for thread in threads:
        head = f"{indent}shapes.add_comment(slide, author={thread.get('author', '')!r}"
        head += f", created={thread.get('created', '')!r}"
        if thread.get("author_id"):
            head += f", author_id={thread['author_id']!r}"
        if thread.get("resolved"):
            head += ", resolved=True"
        if thread.get("pending"):
            head += ", pending=True"
        lines.append(head + ",")
        lines.append(f"{indent}    text={thread.get('text', '')!r}")
        replies = thread.get("replies") or []
        if replies:
            lines[-1] += ","
            lines.append(f"{indent}    replies=[")
            for reply in replies:
                lines.append(f"{indent}        {_reply_literal(reply)},")
            lines.append(f"{indent}    ])")
        else:
            lines[-1] += ")"
    lines.append(f"{indent}{REGION_END}")
    return "\n".join(lines)


def parse_comment_calls(source: str) -> list:
    """Read the ``shapes.add_comment(...)`` calls in a slide file back out.

    Used by the deck -> code mirror to spot comments that exist only in the code
    (``pending=True``, added by ``add_comment.py`` and not built yet) so they are
    not mistaken for comments a human deleted in PowerPoint.
    """
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return []
    threads = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if not (
            isinstance(func, ast.Attribute)
            and func.attr == "add_comment"
            and isinstance(func.value, ast.Name)
            and func.value.id == "shapes"
        ):
            continue
        thread = {}
        for kw in node.keywords:
            if kw.arg is None:
                continue
            try:
                thread[kw.arg] = ast.literal_eval(kw.value)
            except (ValueError, SyntaxError):
                pass
        if thread.get("created"):
            threads.append(thread)
    return threads


def _add_slide_end_line(tree: ast.Module):
    """Last line of ``add_slide``'s body, so an inserted region lands inside it."""
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == "add_slide":
            if not node.body:
                return None
            return max(
                getattr(stmt, "end_lineno", stmt.lineno) or stmt.lineno
                for stmt in node.body
            )
    return None


def patch_comment_region(source: str, region: str):
    """Replace, insert, or remove the comment region in a slide file's source.

    Returns the new source, or ``None`` when nothing needed changing. Raises
    ``ValueError`` if the result would not parse, so callers can keep the old file
    -- a bad comment can then cost a comment, never a slide.
    """
    lines = source.splitlines()
    start = end = None
    for i, line in enumerate(lines):
        if line.strip() == REGION_START:
            start = i
        elif line.strip() == REGION_END and start is not None:
            end = i
            break

    if start is not None and end is not None:
        new_lines = lines[:start] + (region.splitlines() if region else []) + lines[end + 1:]
    elif not region:
        return None  # no region present, none wanted
    else:
        try:
            tree = ast.parse(source)
        except SyntaxError as exc:
            raise ValueError(f"slide file does not parse: {exc}") from exc
        anchor = _add_slide_end_line(tree)
        if anchor is None:
            raise ValueError("no add_slide() body to attach comments to")
        new_lines = lines[:anchor] + region.splitlines() + lines[anchor:]

    new_source = "\n".join(new_lines)
    if source.endswith("\n"):
        new_source += "\n"
    if new_source == source:
        return None
    try:
        ast.parse(new_source)
    except SyntaxError as exc:
        raise ValueError(f"patched slide file would not parse: {exc}") from exc
    return new_source
