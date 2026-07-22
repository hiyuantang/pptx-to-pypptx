"""Extract PowerPoint modern threaded comments from a target deck.

python-pptx cannot represent comment parts, so scaffold.py copies them verbatim
into the project's ``comments/`` directory and they are re-attached after each
build (see ``template/lib/comments.py``). This module only reads the source
package; it never mutates it.

The manifest maps the 1-based deck slide index (``slide*.xml`` filename order,
matching generate_slides.py) to the comment part filenames attached to that
slide, so the injector can re-wire each thread to the rebuilt slide.
"""

import json
import re
import shutil
import zipfile
from pathlib import Path

# Slide -> comment-thread relationship used by modern (2018) PowerPoint comments.
COMMENTS_REL_TYPE = "http://schemas.microsoft.com/office/2018/10/relationships/comments"


def _slide_key(name: str) -> int:
    m = re.search(r"(\d+)", name)
    return int(m.group(1)) if m else 0


def extract_comments(target: Path, output_dir: Path) -> int:
    """Copy authors + per-slide comment parts into ``<output_dir>/comments/``.

    Returns the number of slides that carry comments (``0`` if the deck has
    none, in which case no ``comments/`` directory is written). Only comment
    parts actually referenced by a slide are copied.
    """
    target = Path(target)
    comments_out = Path(output_dir) / "comments"
    with zipfile.ZipFile(target, "r") as zf:
        names = set(zf.namelist())
        if not any(re.match(r"ppt/comments/.*\.xml$", n) for n in names):
            return 0  # no comment parts -> nothing to preserve

        slide_files = sorted(
            (n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=_slide_key,
        )
        slides_map: dict[str, list[str]] = {}
        referenced: set[str] = set()
        for idx, slide_name in enumerate(slide_files, start=1):
            rels_name = f"ppt/slides/_rels/{slide_name.split('/')[-1]}.rels"
            if rels_name not in names:
                continue
            rels_xml = zf.read(rels_name).decode("utf-8")
            targets = []
            for tag in re.findall(r"<Relationship\b[^>]*/>", rels_xml):
                attrs = dict(re.findall(r'(\w+)="([^"]*)"', tag))
                if attrs.get("Type") == COMMENTS_REL_TYPE and attrs.get("Target"):
                    fname = attrs["Target"].split("/")[-1]
                    targets.append(fname)
                    referenced.add(fname)
            if targets:
                slides_map[str(idx)] = targets

        if not slides_map:
            return 0

        if comments_out.exists():
            shutil.rmtree(comments_out)
        comments_out.mkdir(parents=True)

        for fname in sorted(referenced):
            part = f"ppt/comments/{fname}"
            if part in names:
                (comments_out / fname).write_bytes(zf.read(part))

        manifest: dict = {"slides": slides_map}
        if "ppt/authors.xml" in names:
            (comments_out / "authors.xml").write_bytes(zf.read("ppt/authors.xml"))
            manifest["authors"] = "authors.xml"

    (comments_out / "manifest.json").write_text(
        json.dumps(manifest, indent=2), encoding="utf-8"
    )
    return len(slides_map)
