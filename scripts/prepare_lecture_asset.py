#!/usr/bin/env python3
"""Prepare one Markdown-ready PNG from an image or rendered slide region."""

import argparse
import tempfile
from pathlib import Path

from PIL import Image, UnidentifiedImageError

from helpers.lecture_assets import parse_hex_color, remove_edge_background, trim_transparent
from helpers.pptx_utils import count_slides, read_slide_size
from helpers.rendering import render_slides


def parse_bounds(value: str) -> tuple[float, float, float, float]:
    try:
        parts = tuple(float(part.strip()) for part in value.split(","))
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Bounds must be x,y,w,h in slide inches") from exc
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("Bounds must contain four values: x,y,w,h")
    x, y, width, height = parts
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        raise argparse.ArgumentTypeError("Bounds must have non-negative x/y and positive width/height")
    return parts


def crop_slide_region(
    image: Image.Image,
    bounds: tuple[float, float, float, float],
    slide_size: tuple[float, float],
) -> Image.Image:
    x, y, width, height = bounds
    slide_width, slide_height = slide_size
    if x + width > slide_width + 1e-6 or y + height > slide_height + 1e-6:
        raise ValueError(
            f"Bounds {bounds} exceed the {slide_width:.3f}×{slide_height:.3f} inch slide canvas"
        )
    left = round(x / slide_width * image.width)
    top = round(y / slide_height * image.height)
    right = round((x + width) / slide_width * image.width)
    bottom = round((y + height) / slide_height * image.height)
    if right <= left or bottom <= top:
        raise ValueError("Bounds collapse to an empty pixel region at the selected DPI")
    return image.crop((left, top, right, bottom))


def load_source_image(
    *,
    input_image: Path | None,
    target: Path | None,
    slide: int | None,
    bounds: tuple[float, float, float, float] | None,
    dpi: int,
) -> Image.Image:
    if input_image is not None:
        try:
            with Image.open(input_image) as image:
                if bool(getattr(image, "is_animated", False)):
                    raise ValueError("Animated images must be preserved as-is, not flattened by this helper")
                return image.convert("RGBA")
        except UnidentifiedImageError as exc:
            raise ValueError(f"Unsupported input image: {input_image}") from exc

    if target is None or slide is None:
        raise ValueError("A PPTX target requires --slide")
    total = count_slides(target)
    if not 1 <= slide <= total:
        raise ValueError(f"Slide {slide} is outside the deck's 1-{total} range")
    with tempfile.TemporaryDirectory() as tmpdir:
        rendered = render_slides(target, [slide], Path(tmpdir), dpi)[slide]
        with Image.open(rendered) as image:
            result = image.convert("RGBA")
    if bounds is not None:
        result = crop_slide_region(result, bounds, read_slide_size(target))
    return result


def prepare_asset(
    image: Image.Image,
    *,
    transparent: bool,
    background: tuple[int, int, int] | None,
    tolerance: int,
    trim: bool,
    padding: int,
) -> tuple[Image.Image, dict | None]:
    stats = None
    result = image.convert("RGBA")
    if transparent:
        result, stats = remove_edge_background(
            result,
            tolerance=tolerance,
            background=background,
        )
    if trim:
        result = trim_transparent(result, padding)
    return result, stats


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Prepare a Markdown-ready PNG from an existing image or a tightly "
            "bounded PPTX slide region."
        )
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--input-image", help="Raster source image")
    source.add_argument("--target", help="Source PPTX to render")
    parser.add_argument("--slide", type=int, help="1-based slide number for --target")
    parser.add_argument("--bounds", type=parse_bounds, help="Optional x,y,w,h crop in slide inches")
    parser.add_argument("--output", required=True, help="Output PNG")
    parser.add_argument("--dpi", type=int, default=300, help="PPTX render DPI (default: 300)")
    parser.add_argument(
        "--transparent",
        action="store_true",
        help="Remove only edge-connected pixels near a flat background color",
    )
    parser.add_argument("--background", help="Explicit background color as #RGB or #RRGGBB")
    parser.add_argument(
        "--tolerance",
        type=int,
        default=12,
        help="Maximum per-channel background difference, 0-255 (default: 12)",
    )
    parser.add_argument("--trim", action="store_true", help="Trim transparent outer pixels")
    parser.add_argument("--padding", type=int, default=0, help="Transparent padding added after --trim")
    parser.add_argument("--overwrite", action="store_true", help="Replace an existing output file")
    args = parser.parse_args()

    input_image = Path(args.input_image) if args.input_image else None
    target = Path(args.target) if args.target else None
    output = Path(args.output)
    if input_image and (args.slide is not None or args.bounds is not None):
        parser.error("--slide and --bounds apply only to --target")
    if target and args.slide is None:
        parser.error("--target requires --slide")
    if args.background and not args.transparent:
        parser.error("--background requires --transparent")
    if args.padding and not args.trim:
        parser.error("--padding requires --trim")
    if output.suffix.lower() != ".png":
        parser.error("--output must end in .png")
    if output.exists() and not args.overwrite:
        raise FileExistsError(f"Output already exists: {output}")
    if args.dpi <= 0:
        parser.error("--dpi must be positive")

    image = load_source_image(
        input_image=input_image,
        target=target,
        slide=args.slide,
        bounds=args.bounds,
        dpi=args.dpi,
    )
    background = parse_hex_color(args.background) if args.background else None
    prepared, stats = prepare_asset(
        image,
        transparent=args.transparent,
        background=background,
        tolerance=args.tolerance,
        trim=args.trim,
        padding=args.padding,
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    prepared.save(output, format="PNG")
    print(f"Wrote lecture asset to {output} ({prepared.width}×{prepared.height}px)")
    if stats:
        print(
            f"Removed {stats['removed_fraction']:.1%} edge-connected background "
            f"using {stats['background']} at tolerance {stats['tolerance']}."
        )
    print("Inspect the output visually before linking it from lecture notes.")


if __name__ == "__main__":
    main()
