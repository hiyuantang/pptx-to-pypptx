"""Image inspection and conservative background-removal helpers."""

import hashlib
from collections import deque
from pathlib import Path

from PIL import Image


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def inspect_raster(path: Path) -> dict:
    """Return Markdown-relevant raster metadata without changing the image."""
    with Image.open(path) as image:
        image.seek(0)
        rgba = image.convert("RGBA")
        alpha_min, alpha_max = rgba.getchannel("A").getextrema()
        return {
            "format": image.format,
            "width_px": image.width,
            "height_px": image.height,
            "mode": image.mode,
            "animated": bool(getattr(image, "is_animated", False)),
            "frame_count": int(getattr(image, "n_frames", 1)),
            "has_transparency": alpha_min < 255,
            "alpha_range": [alpha_min, alpha_max],
        }


def difference_hash(path: Path) -> str | None:
    """Return a 64-bit dHash for comparing raster candidates, or None on failure."""
    try:
        with Image.open(path) as image:
            image.seek(0)
            resampling = getattr(Image, "Resampling", Image)
            gray = image.convert("RGBA").convert("L").resize((9, 8), resampling.LANCZOS)
            pixel_data = gray.get_flattened_data() if hasattr(gray, "get_flattened_data") else gray.getdata()
            pixels = list(pixel_data)
    except Exception:
        return None

    bits = 0
    for row in range(8):
        offset = row * 9
        for column in range(8):
            bits = (bits << 1) | int(pixels[offset + column] > pixels[offset + column + 1])
    return f"{bits:016x}"


def parse_hex_color(value: str) -> tuple[int, int, int]:
    raw = value.strip().lstrip("#")
    if len(raw) == 3:
        raw = "".join(character * 2 for character in raw)
    if len(raw) != 6:
        raise ValueError("Background color must be #RGB or #RRGGBB")
    try:
        return tuple(int(raw[index:index + 2], 16) for index in (0, 2, 4))
    except ValueError as exc:
        raise ValueError("Background color must be hexadecimal") from exc


def _color_distance(left: tuple[int, ...], right: tuple[int, ...]) -> int:
    return max(abs(left[index] - right[index]) for index in range(3))


def _infer_corner_background(image: Image.Image, tolerance: int) -> tuple[int, int, int]:
    width, height = image.size
    pixels = image.load()
    corner_pixels = [
        pixels[0, 0],
        pixels[width - 1, 0],
        pixels[0, height - 1],
        pixels[width - 1, height - 1],
    ]
    corners = [pixel[:3] for pixel in corner_pixels if pixel[3] == 255]
    if len(corners) < 3:
        raise ValueError(
            "Cannot infer a flat background from fewer than three opaque corners; "
            "the image may already be transparent, so pass --background explicitly "
            "or omit --transparent"
        )
    best_cluster = []
    for candidate in corners:
        cluster = [color for color in corners if _color_distance(candidate, color) <= tolerance]
        if len(cluster) > len(best_cluster):
            best_cluster = cluster
    if len(best_cluster) < 3:
        raise ValueError(
            "Cannot infer a uniform edge background from at least three corners; "
            "pass --background explicitly or keep the render opaque"
        )
    channels = zip(*best_cluster)
    return tuple(round(sum(values) / len(best_cluster)) for values in channels)


def remove_edge_background(
    image: Image.Image,
    *,
    tolerance: int = 12,
    background: tuple[int, int, int] | None = None,
) -> tuple[Image.Image, dict]:
    """Make only edge-connected pixels near a flat background color transparent.

    Enclosed pixels of the same color remain opaque, so white equation fills and
    labels surrounded by artwork are not globally keyed out.
    """
    if not 0 <= tolerance <= 255:
        raise ValueError("Tolerance must be between 0 and 255")

    rgba = image.convert("RGBA")
    width, height = rgba.size
    if width == 0 or height == 0:
        raise ValueError("Image is empty")
    background = background or _infer_corner_background(rgba, tolerance)
    pixels = rgba.load()
    visited = bytearray(width * height)
    queue: deque[tuple[int, int]] = deque()

    def index(x: int, y: int) -> int:
        return y * width + x

    def matches(x: int, y: int) -> bool:
        pixel = pixels[x, y]
        return pixel[3] == 0 or _color_distance(pixel, background) <= tolerance

    def enqueue(x: int, y: int) -> None:
        position = index(x, y)
        if not visited[position] and matches(x, y):
            visited[position] = 1
            queue.append((x, y))

    for x in range(width):
        enqueue(x, 0)
        enqueue(x, height - 1)
    for y in range(height):
        enqueue(0, y)
        enqueue(width - 1, y)

    removed = 0
    while queue:
        x, y = queue.popleft()
        red, green, blue, alpha = pixels[x, y]
        if alpha:
            pixels[x, y] = (red, green, blue, 0)
            removed += 1
        if x:
            enqueue(x - 1, y)
        if x + 1 < width:
            enqueue(x + 1, y)
        if y:
            enqueue(x, y - 1)
        if y + 1 < height:
            enqueue(x, y + 1)

    stats = {
        "background": "#" + "".join(f"{channel:02x}" for channel in background),
        "tolerance": tolerance,
        "removed_pixels": removed,
        "removed_fraction": removed / (width * height),
    }
    return rgba, stats


def recover_alpha_from_mattes(dark: Image.Image, light: Image.Image) -> Image.Image:
    """Recover transparent artwork from matching black- and white-matte renders.

    A single chroma-key render leaves colored fringes on antialiased PowerPoint
    shapes. Two renders provide enough information to reconstruct both alpha and
    the foreground color: dark = alpha * foreground, while
    light = dark + (1 - alpha) * 255.
    """
    if dark.size != light.size:
        raise ValueError("Black- and white-matte renders must have matching dimensions")

    dark_rgb = dark.convert("RGB")
    light_rgb = light.convert("RGB")
    result = Image.new("RGBA", dark.size)
    # Work in strips so a 300-DPI widescreen slide does not allocate a second
    # full-frame Python tuple list (millions of pixels) while reconstructing.
    strip_height = 64
    for top in range(0, dark.height, strip_height):
        bottom = min(dark.height, top + strip_height)
        dark_strip = dark_rgb.crop((0, top, dark.width, bottom))
        light_strip = light_rgb.crop((0, top, light.width, bottom))
        dark_data = (
            dark_strip.get_flattened_data()
            if hasattr(dark_strip, "get_flattened_data")
            else dark_strip.getdata()
        )
        light_data = (
            light_strip.get_flattened_data()
            if hasattr(light_strip, "get_flattened_data")
            else light_strip.getdata()
        )
        recovered = []
        for dark_pixel, light_pixel in zip(dark_data, light_data):
            delta_red = max(0, min(255, light_pixel[0] - dark_pixel[0]))
            delta_green = max(0, min(255, light_pixel[1] - dark_pixel[1]))
            delta_blue = max(0, min(255, light_pixel[2] - dark_pixel[2]))
            median_delta = (
                delta_red + delta_green + delta_blue
                - min(delta_red, delta_green, delta_blue)
                - max(delta_red, delta_green, delta_blue)
            )
            alpha = 255 - median_delta
            if alpha <= 0:
                recovered.append((0, 0, 0, 0))
                continue
            foreground = tuple(
                max(0, min(255, round(channel * 255 / alpha)))
                for channel in dark_pixel
            )
            recovered.append((*foreground, alpha))
        recovered_strip = Image.new("RGBA", dark_strip.size)
        recovered_strip.putdata(recovered)
        result.paste(recovered_strip, (0, top))
    return result


def trim_transparent(image: Image.Image, padding: int = 0) -> Image.Image:
    """Crop to non-transparent content and add transparent pixel padding."""
    if padding < 0:
        raise ValueError("Padding cannot be negative")
    rgba = image.convert("RGBA")
    bounds = rgba.getchannel("A").getbbox()
    if bounds is None:
        raise ValueError("Image contains no visible pixels after background removal")
    cropped = rgba.crop(bounds)
    if not padding:
        return cropped
    canvas = Image.new("RGBA", (cropped.width + 2 * padding, cropped.height + 2 * padding), (0, 0, 0, 0))
    # Paste without a mask so semi-transparent antialiased edge pixels retain
    # their original alpha instead of having it multiplied a second time.
    canvas.paste(cropped, (padding, padding))
    return canvas
