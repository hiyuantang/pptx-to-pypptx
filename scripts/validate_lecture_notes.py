#!/usr/bin/env python3
"""Validate lecture-note image links and local asset hygiene."""

import argparse
import html
import re
import sys
import unicodedata
from collections.abc import Iterable
from pathlib import Path
from urllib.parse import unquote, urlparse

from helpers.lecture_assets import inspect_raster


MARKDOWN_IMAGE_LINK = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
HTML_IMAGE_TAG = re.compile(r"<img\b[^>]*>", re.IGNORECASE | re.DOTALL)
HTML_ATTRIBUTE = re.compile(
    r"(?P<name>[A-Za-z_:][-A-Za-z0-9_:.]*)\s*=\s*"
    r"(?:\"(?P<double>[^\"]*)\"|'(?P<single>[^']*)')"
)
FENCED_CODE = re.compile(r"```.*?```", re.DOTALL)
NONSTANDARD_MATH_DELIMITER = re.compile(r"\\(?:\(|\)|\[|\])")
MARKDOWN_MEDIA = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}
TEMPORARY_SLIDE_IMAGE = re.compile(r"(?i)^slide[_-]?\d+\.(?:png|jpe?g|webp)$")
SOURCE_PROVENANCE = re.compile(
    r"(?mi)<!--[ \t]*(?:lecture-source-slide|source-slides?)[ \t]*:"
)
SLIDE_HEADING = re.compile(r"(?mi)^#{1,6}[ \t]+slide[ \t]+\d+\b")
SOURCE_TITLE = re.compile(r"(?mi)^#[ \t]+speaker notes[ \t]*:")
MARKDOWN_HEADING = re.compile(
    r"(?m)^(#{1,6})[ \t]+(.+?)(?:[ \t]+#+)?[ \t]*$"
)
TOC_LIST_ITEM = re.compile(
    r"(?m)^([ \t]*)[-*+][ \t]+\[([^]]+)\]\(#([^)]+)\)[ \t]*$"
)
VISUAL_LEDGER_MEDIA = re.compile(
    r"(?i)(?<![-A-Za-z0-9_.])([A-Za-z0-9][A-Za-z0-9_.-]*\.(?:png|jpe?g|gif|webp|svg))"
)
VISUAL_LEDGER_RENDER_FAILURE = re.compile(
    r"(?i)\b(?:libreoffice|powerpoint|render(?:er|ing|ed)?|font substitution|"
    r"wrapp?(?:ed|ing|s)?|clipp?(?:ed|ing)?|blank|illegible|layout|geometry|"
    r"extract(?:ion|ed|ing)?|opaque|transparen(?:cy|t))\b"
)
VISUAL_LEDGER_HIGH_RISK = re.compile(
    r"(?i)\b(?:diagram|pipeline|architecture|arrow|flow|process|progressive|"
    r"before[/-]after|before-and-after|workflow|state|model output|annotated|"
    r"task head|head pipeline)\b"
)
VISUAL_LEDGER_TEXT_SUBSTITUTE = re.compile(
    r"(?i)\b(?:markdown|prose|table|list|equation|math|text)\b"
)
VISUAL_LEDGER_NO_SPATIAL = re.compile(
    r"(?i)\bno instructional spatial relationship\b"
)


def _link_target(raw: str) -> str:
    value = raw.strip()
    if value.startswith("<") and ">" in value:
        return value[1:value.index(">")]
    return value.split(maxsplit=1)[0]


def _image_references(text: str) -> tuple[list[tuple[str, str]], list[str]]:
    """Return ``(alt, local target)`` pairs from Markdown and raw HTML images."""

    references = [
        (match.group(1), _link_target(match.group(2)))
        for match in MARKDOWN_IMAGE_LINK.finditer(text)
    ]
    errors: list[str] = []
    for tag_match in HTML_IMAGE_TAG.finditer(text):
        tag = tag_match.group(0)
        attributes: dict[str, str] = {}
        for attribute in HTML_ATTRIBUTE.finditer(tag):
            value = attribute.group("double")
            if value is None:
                value = attribute.group("single")
            attributes[attribute.group("name").lower()] = html.unescape(value)
        source = attributes.get("src")
        if source is None:
            errors.append("Raw HTML image is missing a quoted src attribute")
            continue
        if "alt" not in attributes:
            errors.append(f"Raw HTML image is missing an alt attribute: {source}")
        references.append((attributes.get("alt", ""), source.strip()))
    return references, errors


def _plain_heading_title(value: str) -> str:
    """Return the visible heading text used for a TOC label and anchor."""
    value = re.sub(r"!\[([^]]*)\]\([^)]+\)", r"\1", value)
    value = re.sub(r"\[([^]]+)\]\([^)]+\)", r"\1", value)
    value = re.sub(r"<[^>]+>", "", value)
    value = re.sub(r"[*_`~]", "", value)
    return html.unescape(value).strip()


def _github_heading_slug(title: str, seen: dict[str, int]) -> str:
    """Generate the GitHub-style fragment used by common Markdown renderers."""
    base = "".join(
        character
        for character in title.lower()
        if (
            character in {" ", "-"}
            or character.isalnum()
            or unicodedata.category(character) in {"Mn", "Mc", "Pc"}
        )
    )
    base = re.sub(r"[ \t]+", "-", base.strip())
    duplicate_index = seen.get(base, 0)
    seen[base] = duplicate_index + 1
    return base if duplicate_index == 0 else f"{base}-{duplicate_index}"


def _validate_table_of_contents(prose: str) -> tuple[list[str], int]:
    errors: list[str] = []
    headings = []
    seen_slugs: dict[str, int] = {}
    for match in MARKDOWN_HEADING.finditer(prose):
        level = len(match.group(1))
        title = _plain_heading_title(match.group(2))
        headings.append({
            "level": level,
            "title": title,
            "slug": _github_heading_slug(title, seen_slugs),
            "start": match.start(),
            "end": match.end(),
        })

    toc_headings = [
        heading
        for heading in headings
        if heading["level"] == 2 and heading["title"] == "Table of Contents"
    ]
    if not toc_headings:
        return [
            "Missing required ## Table of Contents after the title/opening paragraph"
        ], 0
    if len(toc_headings) > 1:
        errors.append("Use exactly one ## Table of Contents section")
    toc = toc_headings[0]

    title_headings = [heading for heading in headings if heading["level"] == 1]
    if len(title_headings) != 1:
        errors.append("Use exactly one # lecture title before the Table of Contents")
    elif title_headings[0]["start"] > toc["start"]:
        errors.append("Place the # lecture title before ## Table of Contents")
    else:
        introduction = prose[title_headings[0]["end"]:toc["start"]].strip()
        introduction_blocks = [
            block for block in re.split(r"\n[ \t]*\n", introduction) if block.strip()
        ]
        if len(introduction_blocks) > 1:
            errors.append(
                "Place ## Table of Contents immediately after the title or its single "
                "opening paragraph"
            )

    earlier_content_heading = next(
        (
            heading
            for heading in headings
            if heading["level"] in {2, 3}
            and heading["title"] != "Table of Contents"
            and heading["start"] < toc["start"]
        ),
        None,
    )
    if earlier_content_heading:
        errors.append(
            "Place ## Table of Contents before the first content section, "
            "immediately after the title/opening paragraph"
        )

    next_heading = next(
        (heading for heading in headings if heading["start"] > toc["start"]),
        None,
    )
    toc_end = next_heading["start"] if next_heading else len(prose)
    toc_body = prose[toc["end"]:toc_end]
    entries = [
        {
            "indent": len(match.group(1).replace("\t", "  ")),
            "title": _plain_heading_title(match.group(2)),
            "slug": match.group(3),
        }
        for match in TOC_LIST_ITEM.finditer(toc_body)
    ]
    non_list_content = TOC_LIST_ITEM.sub("", toc_body).strip()
    if non_list_content:
        errors.append("Table of Contents must contain only a nested Markdown link list")

    expected = [
        heading
        for heading in headings
        if heading["level"] in {2, 3} and heading["title"] != "Table of Contents"
    ]
    if not expected:
        errors.append("Table of Contents requires at least one ## content section")
    actual_pairs = [(entry["title"], entry["slug"]) for entry in entries]
    expected_pairs = [(heading["title"], heading["slug"]) for heading in expected]
    if actual_pairs != expected_pairs:
        errors.append(
            "Table of Contents links must match every ## and ### content heading "
            "exactly, in document order, using GitHub-style anchors"
        )
    elif any(
        (heading["level"] == 2 and entry["indent"] != 0)
        or (heading["level"] == 3 and entry["indent"] < 2)
        for heading, entry in zip(expected, entries)
    ):
        errors.append("Nest ### subsection links under their ## section links")

    return errors, len(entries)


def _normalize_allowed_opaque(
    values: Iterable[str],
    *,
    assets_dir: Path,
) -> tuple[set[str], list[str]]:
    normalized: set[str] = set()
    errors: list[str] = []
    for value in values:
        raw = str(value).strip()
        candidate = Path(raw)
        if not raw or candidate.is_absolute() or ".." in candidate.parts:
            errors.append(
                f"Allowed opaque asset must be a path relative to {assets_dir.name}/: {value}"
            )
            continue
        normalized.add(candidate.as_posix())
    return normalized, errors


def _split_visual_ledger_row(line: str) -> list[str] | None:
    stripped = line.strip()
    if not stripped.startswith("|") or not stripped.endswith("|"):
        return None
    return [cell.strip() for cell in stripped[1:-1].split("|")]


def _validate_visual_ledger(
    visual_ledger_path: Path,
    *,
    linked_files: set[Path],
) -> tuple[list[str], list[str], dict[str, int]]:
    errors: list[str] = []
    warnings: list[str] = []
    summary = {"rows": 0, "kept": 0, "omitted": 0, "retry": 0}
    visual_ledger_path = Path(visual_ledger_path)
    if not visual_ledger_path.exists():
        return [f"Visual coverage ledger not found: {visual_ledger_path}"], warnings, summary

    lines = visual_ledger_path.read_text(encoding="utf-8").splitlines()
    header_index = None
    columns: dict[str, int] = {}
    for index, line in enumerate(lines):
        cells = _split_visual_ledger_row(line)
        if not cells:
            continue
        normalized = [re.sub(r"[^a-z]+", " ", cell.lower()).strip() for cell in cells]
        for key in ("candidate", "disposition", "reason"):
            match = next((i for i, value in enumerate(normalized) if key in value), None)
            if match is not None:
                columns[key] = match
        if len(columns) == 3:
            header_index = index
            break
        columns = {}
    if header_index is None:
        return [
            "Visual coverage ledger must contain Candidate visual, Disposition, "
            "and Reason / QA result columns"
        ], warnings, summary

    kept_assets: set[str] = set()
    started = False
    for index in range(header_index + 1, len(lines)):
        cells = _split_visual_ledger_row(lines[index])
        if cells is None:
            if started:
                break
            continue
        if all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells):
            continue
        started = True
        if len(cells) <= max(columns.values()):
            errors.append(f"Visual ledger row {index + 1} has too few columns")
            continue

        candidate = cells[columns["candidate"]]
        disposition = cells[columns["disposition"]].strip(" `").lower()
        reason = cells[columns["reason"]]
        summary["rows"] += 1
        if disposition not in {"keep", "omit", "retry"}:
            errors.append(
                f"Visual ledger row {index + 1} has invalid disposition: {disposition or '(blank)'}"
            )
            continue
        if disposition == "keep":
            summary["kept"] += 1
            row_assets = set(VISUAL_LEDGER_MEDIA.findall(" ".join(cells)))
            if not row_assets:
                errors.append(
                    f"Visual ledger keep row {index + 1} must record its final asset filename"
                )
            kept_assets.update(row_assets)
        elif disposition == "retry":
            summary["retry"] += 1
            errors.append(
                f"Visual ledger row {index + 1} remains retry; resolve it before delivery"
            )
        else:
            summary["omitted"] += 1
            if VISUAL_LEDGER_RENDER_FAILURE.search(reason):
                errors.append(
                    f"Visual ledger omit row {index + 1} cites a rendering or extraction failure; "
                    "keep it as retry and resolve it through the fidelity escalation"
                )
            if (
                VISUAL_LEDGER_HIGH_RISK.search(candidate)
                and VISUAL_LEDGER_TEXT_SUBSTITUTE.search(reason)
                and not VISUAL_LEDGER_NO_SPATIAL.search(reason)
            ):
                errors.append(
                    f"Visual ledger omit row {index + 1} replaces a diagrammatic candidate with "
                    "text without stating why there is no instructional spatial relationship"
                )

    if summary["rows"] == 0:
        errors.append("Visual coverage ledger contains no candidate rows")

    linked_assets = {path.name for path in linked_files}
    for asset in sorted(kept_assets - linked_assets):
        errors.append(f"Visual ledger keep asset is not linked from the Markdown: {asset}")
    for asset in sorted(linked_assets - kept_assets):
        errors.append(f"Linked Markdown asset has no visual ledger keep row: {asset}")
    return errors, warnings, summary


def validate_lecture_notes(
    markdown_path: Path,
    assets_dir: Path,
    *,
    strict_transparency: bool = False,
    allowed_opaque: Iterable[str] = (),
    visual_ledger_path: Path | None = None,
) -> dict:
    markdown_path = Path(markdown_path)
    assets_dir = Path(assets_dir)
    errors: list[str] = []
    warnings: list[str] = []
    linked_files: set[Path] = set()
    opaque_files: set[Path] = set()
    visual_ledger_summary = {"rows": 0, "kept": 0, "omitted": 0, "retry": 0}
    allowed_opaque_paths, allowlist_errors = _normalize_allowed_opaque(
        allowed_opaque,
        assets_dir=assets_dir,
    )
    errors.extend(allowlist_errors)
    if allowed_opaque_paths and not strict_transparency:
        errors.append("--allow-opaque requires --strict-transparency")

    if not markdown_path.exists():
        return {"errors": [f"Markdown file not found: {markdown_path}"], "warnings": [], "summary": {}}
    if not assets_dir.exists():
        errors.append(f"Assets directory not found: {assets_dir}")

    markdown_root = markdown_path.parent.resolve()
    assets_root = assets_dir.resolve()
    text = markdown_path.read_text(encoding="utf-8")
    prose = FENCED_CODE.sub("", text)
    toc_errors, toc_links = _validate_table_of_contents(prose)
    errors.extend(toc_errors)
    if NONSTANDARD_MATH_DELIMITER.search(prose):
        errors.append(
            "Use $...$ and $$...$$ for Markdown math, not \\(...\\) or \\[...\\]"
        )
    if SOURCE_PROVENANCE.search(prose):
        errors.append("Source-slide provenance remains; run finalize_lecture_notes.py")
    if SLIDE_HEADING.search(prose) or SOURCE_TITLE.search(prose):
        errors.append("Slide-by-slide source structure remains; organize the final notes by concept")
    image_references, image_reference_errors = _image_references(prose)
    errors.extend(image_reference_errors)
    if not image_references:
        warnings.append("No lecture-note image links found")

    for raw_alt, raw_target in image_references:
        alt = html.unescape(raw_alt).strip()
        target = unquote(html.unescape(raw_target))
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
    allowed_opaque_used: set[str] = set()
    for path in sorted(opaque_files):
        relative_path = path.relative_to(assets_root).as_posix()
        if strict_transparency:
            if relative_path in allowed_opaque_paths:
                allowed_opaque_used.add(relative_path)
            else:
                errors.append(
                    "Opaque raster asset is not explicitly allowed in strict mode: "
                    f"{relative_path} (use --allow-opaque {relative_path} only for an "
                    "intrinsic screenshot, photo, or panel)"
                )
        else:
            warnings.append(
                "Opaque raster asset (inspect whether transparency is appropriate): "
                f"{path.name}"
            )
    if strict_transparency:
        for relative_path in sorted(allowed_opaque_paths - allowed_opaque_used):
            errors.append(
                f"Allowed opaque asset is not a linked opaque raster: {relative_path}"
            )

    if visual_ledger_path is not None:
        ledger_errors, ledger_warnings, visual_ledger_summary = _validate_visual_ledger(
            visual_ledger_path,
            linked_files=linked_files,
        )
        errors.extend(ledger_errors)
        warnings.extend(ledger_warnings)

    return {
        "errors": errors,
        "warnings": warnings,
        "summary": {
            "image_links": len(image_references),
            "linked_assets": len(linked_files),
            "unused_assets": len(unused),
            "opaque_rasters": len(opaque_files),
            "allowed_opaque_rasters": len(allowed_opaque_used),
            "toc_links": toc_links,
            "visual_ledger_rows": visual_ledger_summary["rows"],
            "visual_ledger_kept": visual_ledger_summary["kept"],
            "visual_ledger_omitted": visual_ledger_summary["omitted"],
            "visual_ledger_retry": visual_ledger_summary["retry"],
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
    parser.add_argument(
        "--strict-transparency",
        action="store_true",
        help="Fail on every opaque raster that is not explicitly allowlisted",
    )
    parser.add_argument(
        "--allow-opaque",
        action="append",
        default=[],
        metavar="ASSET",
        help=(
            "Allow one intentional opaque raster, as a path relative to --assets-dir; "
            "repeat for additional screenshots, photos, or intrinsic panels"
        ),
    )
    parser.add_argument(
        "--visual-ledger",
        required=True,
        help="Scratch visual-coverage ledger used to justify every keep, omit, or retry decision",
    )
    args = parser.parse_args()

    markdown = Path(args.markdown)
    assets_dir = Path(args.assets_dir) if args.assets_dir else markdown.parent / "lecture-notes-assets"
    result = validate_lecture_notes(
        markdown,
        assets_dir,
        strict_transparency=args.strict_transparency,
        allowed_opaque=args.allow_opaque,
        visual_ledger_path=Path(args.visual_ledger),
    )
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
            f"{summary['opaque_rasters']} opaque raster(s), "
            f"{summary['allowed_opaque_rasters']} explicitly allowed, "
            f"{summary['toc_links']} TOC link(s), "
            f"{summary['visual_ledger_rows']} visual-ledger row(s) "
            f"({summary['visual_ledger_kept']} keep, "
            f"{summary['visual_ledger_omitted']} omit, "
            f"{summary['visual_ledger_retry']} retry)."
        )
    if result["errors"]:
        sys.exit(1)
    print("Lecture notes validation passed.")


if __name__ == "__main__":
    main()
