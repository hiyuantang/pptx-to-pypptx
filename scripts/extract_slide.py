#!/usr/bin/env python3
"""Extract shape layout from a target PPTX for one or more slides.

Usage:
    uv run python extract_slide.py "path/to/target.pptx" 1
    uv run python extract_slide.py "path/to/target.pptx" 8-12
    uv run python extract_slide.py "path/to/target.pptx" 4,5,9
    uv run python extract_slide.py "path/to/target.pptx" all
    uv run python extract_slide.py --verbose "path/to/target.pptx" 16
    uv run python extract_slide.py --json "path/to/target.pptx" 7 --screenshot
"""

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from helpers.pptx_utils import count_slides, parse_slide_range
from helpers.slide_xml import read_slide_shapes, parse_slide_hidden


def _font_summary(shape):
    paras = shape.get("paragraphs", [])
    if not paras:
        return None
    for run in paras[0].get("runs", []):
        if run.get("text") in (None, "", "\n"):
            continue
        parts = []
        face = run.get("typeface") or run.get("ea") or run.get("cs")
        if face:
            parts.append(f"font={face}")
        if run.get("sz"):
            parts.append(f"sz={run['sz']}")
        if run.get("color"):
            parts.append(f"color={run['color']}")
        return " ".join(parts) if parts else None
    return None


def _format_shape(shape, verbose=False):
    text = re.sub(r"\s+", " ", (shape.get("text") or "").strip())
    line = f"{shape['type']:12s} x={shape['x']:.3f}, y={shape['y']:.3f}, w={shape['w']:.3f}, h={shape['h']:.3f}"
    if text:
        line += f" | {text[:200]}"
    if not verbose:
        return line

    extras = []
    if shape.get("z") is not None:
        extras.append(f"z={shape['z']}")
    if shape.get("rot") and shape["rot"] != "0":
        extras.append(f"rot={shape['rot']}")
    if shape.get("flipH"):
        extras.append("flipH")
    if shape.get("flipV"):
        extras.append("flipV")
    if shape.get("fill"):
        extras.append(f"fill={shape['fill']}")
    if shape.get("line"):
        extras.append(f"line={shape['line']}")
    if shape.get("anchor"):
        extras.append(f"anchor={shape['anchor']}")
    if shape.get("wrap"):
        extras.append(f"wrap={shape['wrap']}")
    if shape.get("autofit"):
        extras.append(f"autofit={shape['autofit']}")
    if shape.get("effects"):
        extras.append(f"effects={shape['effects']}")
    if shape.get("imgHash"):
        extras.append(f"img={shape.get('imgFile')} hash={shape['imgHash'][:8]}")
    font = _font_summary(shape)
    if font:
        extras.append(font)
    if shape["type"] == "table" and shape.get("cells"):
        extras.append(f"table={shape['rows']}x{shape['cols']}")

    if extras:
        line += f" | {' | '.join(extras)}"
    return line


def _shape_dict(shape, verbose: bool = False):
    d = {
        "type": shape.get("type"),
        "x": shape.get("x"),
        "y": shape.get("y"),
        "w": shape.get("w"),
        "h": shape.get("h"),
        "text": re.sub(r"\s+", " ", (shape.get("text") or "").strip()) or None,
    }
    if verbose:
        d.update({
            "z": shape.get("z"),
            "rot": shape.get("rot"),
            "flipH": shape.get("flipH"),
            "flipV": shape.get("flipV"),
            "fill": shape.get("fill"),
            "line": shape.get("line"),
            "anchor": shape.get("anchor"),
            "wrap": shape.get("wrap"),
            "autofit": shape.get("autofit"),
            "effects": shape.get("effects"),
            "imgFile": shape.get("imgFile"),
            "imgHash": shape.get("imgHash", "")[:8] or None,
            "font": _font_summary(shape),
            "rows": shape.get("rows") if shape.get("type") == "table" else None,
            "cols": shape.get("cols") if shape.get("type") == "table" else None,
        })
        d = {k: v for k, v in d.items() if v is not None}
    return d


def extract_slide(slide_xml: Path, slide_num: int, verbose: bool = False, json_mode: bool = False):
    if not slide_xml.exists():
        if json_mode:
            return {"slide": slide_num, "error": "Slide not found"}
        print(f"Slide {slide_num} not found")
        return None

    shapes = read_slide_shapes(slide_xml)
    shapes.sort(key=lambda s: (s["y"], s["x"]))
    hidden = parse_slide_hidden(slide_xml)
    if json_mode:
        return {
            "slide": slide_num,
            "hidden": hidden,
            "shape_count": len(shapes),
            "shapes": [_shape_dict(s, verbose) for s in shapes],
        }

    hidden_note = " [HIDDEN]" if hidden else ""
    print(f"Slide {slide_num}: {len(shapes)} shape(s){hidden_note}")
    for shape in shapes:
        print(f"  {_format_shape(shape, verbose=verbose)}")
    return None


def _render_screenshots(pptx_path: Path, slides: list[int], out_dir: Path, dpi: int = 150):
    """Render selected slides to PNG via PDF using LibreOffice + pdftoppm.

    Returns a dict mapping slide number to output PNG path.
    """
    if not shutil.which("soffice"):
        print("Error: LibreOffice (soffice) not found; cannot render screenshots", file=sys.stderr)
        sys.exit(1)
    if not shutil.which("pdftoppm"):
        print("Error: pdftoppm not found; cannot render screenshots", file=sys.stderr)
        sys.exit(1)

    out_dir.mkdir(parents=True, exist_ok=True)
    screenshot_paths = {}
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(pptx_path)],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print(f"LibreOffice conversion failed:\n{result.stderr}", file=sys.stderr)
            sys.exit(1)

        pdf_files = list(tmp.glob("*.pdf"))
        if not pdf_files:
            print("Error: PDF conversion produced no file", file=sys.stderr)
            sys.exit(1)
        pdf_path = pdf_files[0]

        for slide_num in slides:
            prefix = tmp / f"slide_{slide_num}"
            subprocess.run(
                ["pdftoppm", "-png", "-f", str(slide_num), "-l", str(slide_num),
                 "-r", str(dpi), str(pdf_path), str(prefix)],
                check=True,
            )
            rendered = list(tmp.glob(f"slide_{slide_num}-*.png"))
            if not rendered:
                print(f"Warning: no PNG rendered for slide {slide_num}", file=sys.stderr)
                continue
            dest = out_dir / f"slide_{slide_num}.png"
            shutil.move(str(rendered[0]), str(dest))
            screenshot_paths[slide_num] = str(dest)
            print(f"Screenshot: {dest}")
    return screenshot_paths


def main():
    parser = argparse.ArgumentParser(description="Extract slide layout")
    parser.add_argument("pptx", help="Target PPTX")
    parser.add_argument("slide", help="Slide number(s): 14 | 8-12 | 4,5,9 | all")
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Show fill, line, font, anchor, wrap, autofit, rotation, z-order, and image hash"
    )
    parser.add_argument(
        "--screenshot", action="store_true",
        help="Render the selected slide(s) to PNG screenshot(s)"
    )
    parser.add_argument(
        "--screenshot-dir", type=Path, default=Path.cwd(),
        help="Directory for screenshots (default: current directory)"
    )
    parser.add_argument(
        "--screenshot-dpi", type=int, default=150,
        help="DPI for rendered screenshots (default: 150)"
    )
    parser.add_argument(
        "--json", action="store_true",
        help="Output machine-readable JSON (includes screenshot paths when --screenshot is used)"
    )
    args = parser.parse_args()

    pptx_path = Path(args.pptx)
    total = count_slides(pptx_path)
    slides = parse_slide_range(args.slide, total)

    results = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(pptx_path, "r") as zf:
            zf.extractall(tmp)
        for slide_num in slides:
            slide_xml = tmp / "ppt" / "slides" / f"slide{slide_num}.xml"
            result = extract_slide(slide_xml, slide_num, verbose=args.verbose, json_mode=args.json)
            if args.json:
                results.append(result)

    screenshot_paths = {}
    if args.screenshot:
        screenshot_paths = _render_screenshots(pptx_path, slides, args.screenshot_dir, args.screenshot_dpi)

    if args.json:
        for result in results:
            slide_num = result.get("slide")
            if slide_num in screenshot_paths:
                result["screenshot"] = screenshot_paths[slide_num]
        print(json.dumps(results, indent=2))


if __name__ == "__main__":
    main()
