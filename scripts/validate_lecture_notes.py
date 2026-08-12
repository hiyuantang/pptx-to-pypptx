#!/usr/bin/env python3
"""Validate Markdown lecture-note image links and local asset hygiene."""

import argparse
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

from helpers.lecture_assets import inspect_raster


IMAGE_LINK = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
FENCED_CODE = re.compile(r"```.*?```", re.DOTALL)
NONSTANDARD_MATH_DELIMITER = re.compile(r"\\(?:\(|\)|\[|\])")
MARKDOWN_MEDIA = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}
TEMPORARY_SLIDE_IMAGE = re.compile(r"(?i)^slide[_-]?\d+\.(?:png|jpe?g|webp)$")
SOURCE_PROVENANCE = re.compile(
    r"(?mi)<!--[ \t]*(?:lecture-source-slide|source-slides?)[ \t]*:"
)
SLIDE_HEADING = re.compile(r"(?mi)^#{1,6}[ \t]+slide[ \t]+\d+\b")
SOURCE_TITLE = re.compile(r"(?mi)^#[ \t]+speaker notes[ \t]*:")


def _link_target(raw: str) -> str:
    value = raw.strip()
    if value.startswith("<") and ">" in value:
        return value[1:value.index(">")]
    return value.split(maxsplit=1)[0]


def validate_lecture_notes(markdown_path: Path, assets_dir: Path) -> dict:
    markdown_path = Path(markdown_path)
    assets_dir = Path(assets_dir)
    errors: list[str] = []
    warnings: list[str] = []
    linked_files: set[Path] = set()
    opaque_files: set[Path] = set()

    if not markdown_path.exists():
        return {"errors": [f"Markdown file not found: {markdown_path}"], "warnings": [], "summary": {}}
    if not assets_dir.exists():
        errors.append(f"Assets directory not found: {assets_dir}")

    markdown_root = markdown_path.parent.resolve()
    assets_root = assets_dir.resolve()
    text = markdown_path.read_text(encoding="utf-8")
    prose = FENCED_CODE.sub("", text)
    if NONSTANDARD_MATH_DELIMITER.search(prose):
        errors.append(
            "Use $...$ and $$...$$ for Markdown math, not \\(...\\) or \\[...\\]"
        )
    if SOURCE_PROVENANCE.search(prose):
        errors.append("Source-slide provenance remains; run finalize_lecture_notes.py")
    if SLIDE_HEADING.search(prose) or SOURCE_TITLE.search(prose):
        errors.append("Slide-by-slide source structure remains; organize the final notes by concept")
    matches = list(IMAGE_LINK.finditer(text))
    if not matches:
        warnings.append("No Markdown image links found")

    for match in matches:
        alt = match.group(1).strip()
        raw_target = _link_target(match.group(2))
        target = unquote(raw_target)
        if not alt:
            errors.append(f"Image link has empty alt text: {raw_target}")
        target_path = Path(target)
        if "slide-images" in target_path.parts or TEMPORARY_SLIDE_IMAGE.fullmatch(target_path.name):
            errors.append(
                f"Temporary full-slide reference leaked into final lecture notes: {raw_target}"
            )
        parsed = urlparse(target)
        if parsed.scheme or target.startswith("//"):
            errors.append(f"Image link must be local, not remote or data-based: {raw_target}")
            continue
        if Path(target).is_absolute():
            errors.append(f"Image link must be relative: {raw_target}")
            continue

        resolved = (markdown_root / target).resolve()
        try:
            resolved.relative_to(assets_root)
        except ValueError:
            errors.append(f"Image link points outside {assets_dir.name}/: {raw_target}")
            continue
        if resolved.suffix.lower() not in MARKDOWN_MEDIA:
            errors.append(f"Image format is not Markdown-safe: {raw_target}")
        if not resolved.exists():
            errors.append(f"Linked image does not exist: {raw_target}")
            continue
        linked_files.add(resolved)

        if resolved.suffix.lower() != ".svg":
            try:
                info = inspect_raster(resolved)
            except Exception as exc:
                errors.append(f"Cannot read linked image {raw_target}: {exc}")
                continue
            if not info["has_transparency"]:
                opaque_files.add(resolved)

    asset_files = {
        path.resolve()
        for path in assets_dir.iterdir()
        if path.is_file() and not path.name.startswith(".")
    } if assets_dir.exists() else set()
    unused = sorted(asset_files - linked_files)
    for path in unused:
        warnings.append(f"Unreferenced asset: {path.name}")
    for path in sorted(opaque_files):
        warnings.append(f"Opaque raster asset (inspect whether transparency is appropriate): {path.name}")

    return {
        "errors": errors,
        "warnings": warnings,
        "summary": {
            "image_links": len(matches),
            "linked_assets": len(linked_files),
            "unused_assets": len(unused),
            "opaque_rasters": len(opaque_files),
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate Markdown lecture-note image links and assets.")
    parser.add_argument("markdown", help="Lecture-notes Markdown file")
    parser.add_argument(
        "--assets-dir",
        default=None,
        help="Assets folder (default: sibling lecture-notes-assets/)",
    )
    args = parser.parse_args()

    markdown = Path(args.markdown)
    assets_dir = Path(args.assets_dir) if args.assets_dir else markdown.parent / "lecture-notes-assets"
    result = validate_lecture_notes(markdown, assets_dir)
    for message in result["errors"]:
        print(f"ERROR: {message}", file=sys.stderr)
    for message in result["warnings"]:
        print(f"WARNING: {message}")
    summary = result["summary"]
    if summary:
        print(
            f"Checked {summary['image_links']} image link(s): "
            f"{summary['linked_assets']} linked asset(s), "
            f"{summary['unused_assets']} unused, "
            f"{summary['opaque_rasters']} opaque raster(s)."
        )
    if result["errors"]:
        sys.exit(1)
    print("Lecture notes validation passed.")


if __name__ == "__main__":
    main()
