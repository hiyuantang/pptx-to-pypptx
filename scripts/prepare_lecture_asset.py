#!/usr/bin/env python3
"""Prepare one Markdown-ready PNG from an image, slide region, or shape selection."""

import argparse
import tempfile
from pathlib import Path

from PIL import Image, UnidentifiedImageError

from helpers.lecture_assets import (
    parse_hex_color,
    recover_alpha_from_mattes,
    remove_edge_background,
    trim_transparent,
)
from helpers.lecture_shapes import (
    isolate_slide_shapes,
    parse_shape_ids,
)
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


def crop_pdf_page_edge(image: Image.Image, dpi: int) -> Image.Image:
    """Drop the thin PDF page-edge artifact produced by some renderers."""
    inset = max(2, round(dpi / 100))
    if image.width <= inset * 2 or image.height <= inset * 2:
        raise ValueError("Rendered slide is too small to crop its page edge")
    return image.crop((inset, inset, image.width - inset, image.height - inset))


def load_source_image(
    *,
    input_image: Path | None,
    target: Path | None,
    slide: int | None,
    bounds: tuple[float, float, float, float] | None,
    shape_ids: list[int] | None,
    dpi: int,
) -> tuple[Image.Image, dict | None]:
    if input_image is not None:
        try:
            with Image.open(input_image) as image:
                if bool(getattr(image, "is_animated", False)):
                    raise ValueError("Animated images must be preserved as-is, not flattened by this helper")
                return image.convert("RGBA"), None
        except UnidentifiedImageError as exc:
            raise ValueError(f"Unsupported input image: {input_image}") from exc

    if target is None or slide is None:
        raise ValueError("A PPTX target requires --slide")
    total = count_slides(target)
    if not 1 <= slide <= total:
        raise ValueError(f"Slide {slide} is outside the deck's 1-{total} range")
    selection_info = None
    with tempfile.TemporaryDirectory() as tmpdir:
        if shape_ids:
            dark_target = Path(tmpdir) / "selected-shapes-dark.pptx"
            light_target = Path(tmpdir) / "selected-shapes-light.pptx"
            selection_info = isolate_slide_shapes(
                target,
                slide,
                shape_ids,
                dark_target,
                background=(0, 0, 0),
            )
            isolate_slide_shapes(
                target,
                slide,
                shape_ids,
                light_target,
                background=(255, 255, 255),
            )
            dark_render = render_slides(
                dark_target, [slide], Path(tmpdir) / "dark-render", dpi
            )[slide]
            light_render = render_slides(
                light_target, [slide], Path(tmpdir) / "light-render", dpi
            )[slide]
            with Image.open(dark_render) as dark_image, Image.open(light_render) as light_image:
                dark_matte = crop_pdf_page_edge(dark_image.convert("RGB"), dpi)
                light_matte = crop_pdf_page_edge(light_image.convert("RGB"), dpi)
                result = recover_alpha_from_mattes(dark_matte, light_matte)
            selection_info["background"] = "black-and-white mattes"
        else:
            rendered = render_slides(target, [slide], Path(tmpdir) / "rendered", dpi)[slide]
            with Image.open(rendered) as image:
                result = image.convert("RGBA")
    if bounds is not None:
        result = crop_slide_region(result, bounds, read_slide_size(target))
    return result, selection_info


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
            "Prepare a Markdown-ready PNG from an existing image, a tightly bounded "
            "PPTX slide region, or selected native PowerPoint shapes."
        )
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--input-image", help="Raster source image")
    source.add_argument("--target", help="Source PPTX to render")
    parser.add_argument("--slide", type=int, help="1-based slide number for --target")
    parser.add_argument("--bounds", type=parse_bounds, help="Optional x,y,w,h crop in slide inches")
    parser.add_argument(
        "--shape-ids",
        type=parse_shape_ids,
        help=(
            "Comma-separated PowerPoint shape IDs to isolate on --slide. Keeps native "
            "shapes, arrows, text, pictures, and their relationships, reconstructs clean "
            "transparency from black/white matte renders, and trims the result"
        ),
    )
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
    if input_image and (args.slide is not None or args.bounds is not None or args.shape_ids):
        parser.error("--slide, --bounds, and --shape-ids apply only to --target")
    if target and args.slide is None:
        parser.error("--target requires --slide")
    if args.bounds and args.shape_ids:
        parser.error("--bounds and --shape-ids are mutually exclusive")
    if args.shape_ids and args.transparent:
        parser.error("--shape-ids reconstructs transparency automatically; omit --transparent")
    if args.shape_ids and args.background:
        parser.error("--background is not used with --shape-ids")
    if args.background and not args.transparent:
        parser.error("--background requires --transparent")
    if args.padding and not (args.trim or args.shape_ids):
        parser.error("--padding requires --trim")
    if output.suffix.lower() != ".png":
        parser.error("--output must end in .png")
    if output.exists() and not args.overwrite:
        raise FileExistsError(f"Output already exists: {output}")
    if args.dpi <= 0:
        parser.error("--dpi must be positive")

    requested_background = parse_hex_color(args.background) if args.background else None
    image, selection_info = load_source_image(
        input_image=input_image,
        target=target,
        slide=args.slide,
        bounds=args.bounds,
        shape_ids=args.shape_ids,
        dpi=args.dpi,
    )
    transparent = args.transparent
    trim = args.trim or bool(args.shape_ids)
    prepared, stats = prepare_asset(
        image,
        transparent=transparent,
        background=requested_background,
        tolerance=args.tolerance,
        trim=trim,
        padding=args.padding,
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    prepared.save(output, format="PNG")
    print(f"Wrote lecture asset to {output} ({prepared.width}×{prepared.height}px)")
    if selection_info:
        selected = ", ".join(
            f"{shape_id} ({name or 'unnamed'})"
            for shape_id, name in zip(
                selection_info["shape_ids"], selection_info["shape_names"]
            )
        )
        print(f"Isolated slide {selection_info['slide']} shape(s): {selected}")
        print("Recovered transparent antialiased edges from black/white matte renders.")
    if stats:
        print(
            f"Removed {stats['removed_fraction']:.1%} edge-connected background "
            f"using {stats['background']} at tolerance {stats['tolerance']}."
        )
    print("Inspect the output visually before linking it from lecture notes.")


if __name__ == "__main__":
    main()
