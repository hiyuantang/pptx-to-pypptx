#!/usr/bin/env python3
"""Asset sync helpers: deduplicate and copy media from a target PPTX."""

import hashlib
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from helpers.slides import sanitize_name

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _unique_dest_name(dest: Path, used: set[str]) -> str:
    if dest.name not in used:
        return dest.name
    stem = dest.stem
    suffix = dest.suffix
    counter = 1
    while True:
        candidate = f"{stem}_{counter}{suffix}"
        if candidate not in used:
            return candidate
        counter += 1


def _file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def _build_semantic_media_names(slides_dir: Path, media_dir: Path) -> dict[str, str]:
    """Map raw media filenames to semantic names from picture descr fields."""
    media_to_name: dict[str, str] = {}
    if not slides_dir.exists():
        return media_to_name

    rels_dir = slides_dir / "_rels"
    for slide_xml in sorted(slides_dir.glob("slide*.xml")):
        rels_file = rels_dir / f"{slide_xml.name}.rels"
        if not rels_file.exists():
            continue
        rels_root = ET.parse(rels_file).getroot()
        rId_to_target = {
            rel.get("Id"): rel.get("Target")
            for rel in rels_root.iter(f"{{{R_RELS}}}Relationship")
        }

        slide_root = ET.parse(slide_xml).getroot()
        for pic in slide_root.iter(f"{{{P}}}pic"):
            cNvPr = pic.find(f"{{{P}}}nvPicPr/{{{P}}}cNvPr")
            descr = cNvPr.get("descr") if cNvPr is not None else ""
            if not descr:
                continue
            semantic_hint = Path(descr).stem or descr
            blip = pic.find(f"{{{P}}}blipFill/{{{A}}}blip")
            if blip is None:
                continue
            rId = blip.get(f"{{{R}}}embed")
            target = rId_to_target.get(rId)
            if not target:
                continue
            basename = Path(target).name
            if basename not in media_to_name:
                media_to_name[basename] = sanitize_name(semantic_hint)
    return media_to_name


def sync_assets(target: Path, project_dir: Path) -> dict[str, str]:
    """Copy unique media files from target PPTX into project assets/.

    Deduplicates by content hash. Returns a mapping from raw media filename
    (e.g. ``image1.png``) to the asset filename stored in ``assets/``.
    """
    assets_dir = project_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    map_path = assets_dir / "_media_map.json"

    existing_map: dict[str, str] = {}
    if map_path.exists():
        try:
            existing_map = json.loads(map_path.read_text(encoding="utf-8"))
        except Exception:
            existing_map = {}

    # Index existing assets by content hash.
    hash_to_asset: dict[str, str] = {}
    for f in assets_dir.iterdir():
        if f.is_file() and f.name != "_media_map.json":
            hash_to_asset[_file_hash(f)] = f.name

    used_names: set[str] = set(existing_map.values()) | set(hash_to_asset.values())

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        with zipfile.ZipFile(target, "r") as zf:
            zf.extractall(tmp_path)

        slide_dir = tmp_path / "ppt" / "slides"
        media_dir = tmp_path / "ppt" / "media"
        semantic_names = _build_semantic_media_names(slide_dir, media_dir)

        new_map = dict(existing_map)
        if media_dir.exists():
            for f in sorted(media_dir.iterdir()):
                h = _file_hash(f)
                if h in hash_to_asset:
                    new_map[f.name] = hash_to_asset[h]
                    continue

                semantic = semantic_names.get(f.name)
                dest_name = f"{semantic}{f.suffix}" if semantic else f.name
                dest = assets_dir / _unique_dest_name(assets_dir / dest_name, used_names)
                used_names.add(dest.name)
                shutil.copy2(f, dest)
                hash_to_asset[h] = dest.name
                new_map[f.name] = dest.name

    map_path.write_text(json.dumps(new_map, indent=2), encoding="utf-8")
    return new_map


def load_media_map(project_dir: Path) -> dict[str, str]:
    """Load the raw-filename -> asset-filename mapping written by ``sync_assets``."""
    map_path = Path(project_dir) / "assets" / "_media_map.json"
    if not map_path.exists():
        return {}
    try:
        return json.loads(map_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
