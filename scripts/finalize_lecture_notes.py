#!/usr/bin/env python3
"""Remove source-only slide provenance from edited lecture-note Markdown."""

import argparse
import re
from pathlib import Path


PREVIEW_BLOCK = re.compile(
    r"(?ms)^[ \t]*<!-- lecture-source-preview:start -->[ \t]*\n"
    r".*?^[ \t]*<!-- lecture-source-preview:end -->[ \t]*(?:\n|$)"
)
PROVENANCE_LINE = re.compile(
    r"(?m)^[ \t]*<!--[ \t]*(?:lecture-source-slide|source-slides?)[ \t]*:[^>]*-->[ \t]*(?:\n|$)"
)
SLIDE_HEADING = re.compile(r"(?mi)^#{1,6}[ \t]+slide[ \t]+\d+\b[^\n]*$")
SOURCE_TITLE = re.compile(r"(?mi)^#[ \t]+speaker notes[ \t]*:")
TEMPORARY_PREVIEW = re.compile(
    r"(?mi)^!\[[^\]]*temporary[^\]]*slide[^\]]*\]\([^)]+\)[ \t]*$"
)


def finalize_markdown(text: str) -> tuple[str, dict]:
    """Strip owned source markers and reject an unedited slide-by-slide draft."""
    preview_blocks = len(PREVIEW_BLOCK.findall(text))
    provenance_markers = len(PROVENANCE_LINE.findall(text))
    cleaned = PREVIEW_BLOCK.sub("", text)
    cleaned = PROVENANCE_LINE.sub("", cleaned)

    remaining_headings = SLIDE_HEADING.findall(cleaned)
    if remaining_headings:
        examples = "; ".join(heading.strip() for heading in remaining_headings[:3])
        raise ValueError(
            "Draft still contains slide-number headings. Reorganize the prose under concept "
            f"headings before finalizing (for example: {examples})."
        )
    if SOURCE_TITLE.search(cleaned):
        raise ValueError("Replace the source '# Speaker Notes:' title with the lecture title before finalizing")
    if TEMPORARY_PREVIEW.search(cleaned):
        raise ValueError("Draft still contains an unmarked temporary full-slide preview")

    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned).strip() + "\n"
    return cleaned, {
        "preview_blocks_removed": preview_blocks,
        "provenance_markers_removed": provenance_markers,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Remove temporary slide previews and source-slide provenance from an already "
            "edited, concept-organized lecture-note draft."
        )
    )
    parser.add_argument("input", help="Edited lecture-note draft Markdown")
    parser.add_argument("--output", required=True, help="Final lecture-notes Markdown path")
    parser.add_argument("--overwrite", action="store_true", help="Replace an existing output file")
    args = parser.parse_args()

    source = Path(args.input)
    output = Path(args.output)
    if not source.exists():
        raise FileNotFoundError(f"Draft Markdown not found: {source}")
    if source.resolve() == output.resolve() and not args.overwrite:
        parser.error("Refusing to overwrite the input without --overwrite")
    if output.exists() and not args.overwrite:
        raise FileExistsError(f"Output already exists: {output}")

    finalized, stats = finalize_markdown(source.read_text(encoding="utf-8"))
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(finalized, encoding="utf-8")
    print(
        f"Wrote finalized lecture notes to {output}; removed "
        f"{stats['preview_blocks_removed']} preview block(s) and "
        f"{stats['provenance_markers_removed']} provenance marker(s)."
    )


if __name__ == "__main__":
    main()
