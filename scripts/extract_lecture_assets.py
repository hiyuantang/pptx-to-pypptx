#!/usr/bin/env python3
"""Extract embedded lecture-visual candidates and write a provenance manifest.

This script preserves original media whenever it is directly usable in
Markdown. It inventories crop/rotation/luminance transforms and composite
objects that should instead be rendered with ``prepare_lecture_asset.py``.
It deliberately does not decide which progressive state has teaching value.
"""

import argparse
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from PIL import Image, UnidentifiedImageError

from helpers.lecture_assets import difference_hash, inspect_raster, sha256_file
from helpers.pptx_utils import count_slides, parse_slide_range
from helpers.slide_meta import sanitize_name
from helpers.slide_xml import read_slide_shapes


MARKDOWN_MEDIA = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}
GENERIC_SHAPE_NAME = re.compile(r"^(picture|image)(?:[_\s-]*\d+)?$", re.IGNORECASE)


def _bounds(shape: dict) -> dict | None:
    values = [shape.get(key) for key in ("x", "y", "w", "h")]
    if any(value is None for value in values):
        return None
    return {key: round(float(shape[key]), 4) for key in ("x", "y", "w", "h")}


def _walk_images(shapes: list[dict], groups: list[dict] | None = None):
    groups = groups or []
    for shape in shapes:
        if shape.get("type") == "group":
            group = {
                "name": shape.get("name"),
                "id": shape.get("id"),
                "bounds_inches": _bounds(shape),
            }
            yield from _walk_images(shape.get("children") or [], groups + [group])
        elif shape.get("type") == "image":
            yield shape, groups


def _semantic_stem(shape: dict, source_media: str) -> str:
    description = (shape.get("descr") or "").strip()
    if description:
        candidate = sanitize_name(Path(description).stem).replace("_", "-")
        if candidate:
            return candidate
    shape_name = (shape.get("name") or "").strip()
    if shape_name and not GENERIC_SHAPE_NAME.fullmatch(shape_name):
        candidate = sanitize_name(shape_name).replace("_", "-")
        if candidate:
            return candidate
    candidate = sanitize_name(Path(source_media).stem).replace("_", "-")
    return candidate or "lecture-visual"


def _unique_filename(stem: str, suffix: str, used: set[str]) -> str:
    candidate = f"{stem}{suffix.lower()}"
    counter = 2
    while candidate in used:
        candidate = f"{stem}-{counter}{suffix.lower()}"
        counter += 1
    used.add(candidate)
    return candidate


def _save_markdown_media(source: Path, output_dir: Path, stem: str, used: set[str]) -> tuple[Path, bool]:
    suffix = source.suffix.lower()
    if suffix in MARKDOWN_MEDIA:
        destination = output_dir / _unique_filename(stem, suffix, used)
        shutil.copy2(source, destination)
        return destination, False

    try:
        with Image.open(source) as image:
            image.seek(0)
            if bool(getattr(image, "is_animated", False)):
                raise ValueError(f"Animated {suffix or 'image'} is not safely convertible to PNG")
            rgba = image.convert("RGBA")
            destination = output_dir / _unique_filename(stem, ".png", used)
            rgba.save(destination, format="PNG")
            return destination, True
    except (UnidentifiedImageError, OSError) as exc:
        raise ValueError(f"Unsupported media format {suffix or '(none)'}") from exc


def _usage_record(slide_num: int, shape: dict, groups: list[dict], preferred_media: str) -> dict:
    image_info = shape.get("imgInfo") or {}
    render_reasons = []
    if image_info.get("crop"):
        render_reasons.append("source image is cropped on the slide")
    if image_info.get("lum"):
        render_reasons.append("source image has slide luminance adjustments")
    if shape.get("rot") and str(shape.get("rot")) != "0":
        render_reasons.append("source image is rotated on the slide")
    if shape.get("flipH") or shape.get("flipV"):
        render_reasons.append("source image is flipped on the slide")
    if groups:
        render_reasons.append("source image is part of a composite group")

    return {
        "slide": slide_num,
        "shape_id": shape.get("id"),
        "shape_name": shape.get("name"),
        "description": shape.get("descr") or None,
        "source_media": preferred_media,
        "bounds_inches": _bounds(shape),
        "render_bounds_inches": groups[0]["bounds_inches"] if groups else _bounds(shape),
        "groups": groups,
        "crop": image_info.get("crop"),
        "luminance": image_info.get("lum"),
        "rotation": shape.get("rot"),
        "flip_h": bool(shape.get("flipH")),
        "flip_v": bool(shape.get("flipV")),
        "render_recommended": bool(render_reasons),
        "render_reasons": render_reasons,
    }


def _composite_candidates(slide_num: int, shapes: list[dict]) -> list[dict]:
    candidates = []
    for shape in shapes:
        if shape.get("type") not in {"group", "chart", "diagram", "table", "raw"}:
            continue
        candidates.append(
            {
                "slide": slide_num,
                "type": shape.get("type"),
                "shape_id": shape.get("id"),
                "shape_name": shape.get("name"),
                "bounds_inches": _bounds(shape),
                "text_hint": re.sub(r"\s+", " ", (shape.get("text") or "")).strip()[:160] or None,
            }
        )
    return candidates


def extract_lecture_assets(target: Path, output_dir: Path, slides: list[int]) -> dict:
    """Copy Markdown-compatible embedded media and return a usage manifest."""
    target = Path(target)
    output_dir = Path(output_dir)
    if not target.exists():
        raise FileNotFoundError(f"Target PPTX not found: {target}")
    if output_dir.exists() and any(output_dir.iterdir()):
        raise FileExistsError(f"Output directory is not empty: {output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    assets_by_hash: dict[str, dict] = {}
    used_names: set[str] = set()
    unresolved = []
    composites = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(target, "r") as archive:
            archive.extractall(tmp)

        for slide_num in slides:
            slide_xml = tmp / "ppt" / "slides" / f"slide{slide_num}.xml"
            if not slide_xml.exists():
                unresolved.append({"slide": slide_num, "reason": "slide XML is missing"})
                continue
            shapes = read_slide_shapes(slide_xml)
            composites.extend(_composite_candidates(slide_num, shapes))
            for shape, groups in _walk_images(shapes):
                preferred_media = shape.get("svgFile") or shape.get("imgFile")
                if not preferred_media:
                    unresolved.append(
                        {
                            "slide": slide_num,
                            "shape_name": shape.get("name"),
                            "reason": "picture has no resolvable embedded media",
                        }
                    )
                    continue
                source = tmp / "ppt" / "media" / preferred_media
                usage = _usage_record(slide_num, shape, groups, preferred_media)
                if not source.exists():
                    unresolved.append({**usage, "reason": "embedded media file is missing"})
                    continue

                digest = sha256_file(source)
                existing = assets_by_hash.get(digest)
                if existing is not None:
                    if preferred_media not in existing["source_media"]:
                        existing["source_media"].append(preferred_media)
                    existing["usages"].append(usage)
                    continue

                stem = _semantic_stem(shape, preferred_media)
                try:
                    destination, converted = _save_markdown_media(source, output_dir, stem, used_names)
                except ValueError as exc:
                    unresolved.append(
                        {
                            **usage,
                            "reason": f"{exc}; render this visual from its slide region instead",
                        }
                    )
                    continue

                if destination.suffix.lower() == ".svg":
                    media_info = {
                        "format": "SVG",
                        "width_px": None,
                        "height_px": None,
                        "mode": None,
                        "animated": False,
                        "frame_count": 1,
                        "has_transparency": None,
                        "alpha_range": None,
                    }
                    dhash = None
                else:
                    media_info = inspect_raster(destination)
                    dhash = difference_hash(destination)

                assets_by_hash[digest] = {
                    "file": destination.name,
                    "source_media": [preferred_media],
                    "sha256": digest,
                    "difference_hash": dhash,
                    "converted_to_png": converted,
                    **media_info,
                    "usages": [usage],
                }

    assets = sorted(assets_by_hash.values(), key=lambda item: item["file"])
    for asset in assets:
        asset["source_media"].sort()
        asset["usages"].sort(key=lambda usage: (usage["slide"], usage.get("shape_id") or ""))
    exact_repeats = [
        {"file": asset["file"], "slides": sorted({usage["slide"] for usage in asset["usages"]})}
        for asset in assets
        if len(asset["usages"]) > 1
    ]
    return {
        "source": str(target.resolve()),
        "selected_slides": slides,
        "assets": assets,
        "exact_repeat_candidates": exact_repeats,
        "composite_render_candidates": composites,
        "unresolved": unresolved,
        "summary": {
            "extracted_assets": len(assets),
            "resolved_picture_usages": sum(len(asset["usages"]) for asset in assets),
            "total_picture_usages": sum(len(asset["usages"]) for asset in assets) + len(unresolved),
            "exact_repeat_groups": len(exact_repeats),
            "composite_render_candidates": len(composites),
            "unresolved_usages": len(unresolved),
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract embedded lecture-visual candidates and write a provenance manifest."
    )
    parser.add_argument("target", help="Source PPTX")
    parser.add_argument("--output-dir", required=True, help="Empty directory for extracted candidates")
    parser.add_argument("--slides", default="all", help="Slides to inspect: all | 4 | 2-5 | 3,7,9")
    parser.add_argument(
        "--manifest",
        default=None,
        help="Manifest JSON path (default: sibling <output-dir>-manifest.json)",
    )
    args = parser.parse_args()

    target = Path(args.target)
    output_dir = Path(args.output_dir)
    slides = parse_slide_range(args.slides, count_slides(target))
    manifest_path = (
        Path(args.manifest)
        if args.manifest
        else output_dir.parent / f"{output_dir.name}-manifest.json"
    )
    if manifest_path.exists():
        raise FileExistsError(f"Manifest already exists: {manifest_path}")

    manifest = extract_lecture_assets(target, output_dir, slides)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    summary = manifest["summary"]
    print(
        f"Extracted {summary['extracted_assets']} asset(s) from "
        f"{summary['resolved_picture_usages']} resolved picture usage(s) to {output_dir}"
    )
    print(f"Wrote lecture-asset manifest to {manifest_path}")
    if summary["unresolved_usages"]:
        print(
            f"Manifest flags {summary['unresolved_usages']} unresolved usage(s); "
            "render those from slide regions instead."
        )


if __name__ == "__main__":
    main()
