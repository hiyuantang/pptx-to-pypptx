#!/usr/bin/env python3
"""Generate python-pptx slide code from an extracted slide XML.

This is used by generate_slides.py to turn the target PPTX layout into working
slide files instead of empty stubs. It emits code for primitives the
template lib/shapes.py can render; everything else becomes a TODO comment.
"""

import hashlib
import re
from pathlib import Path

from helpers.slide_xml import read_slide_shapes, parse_background, parse_slide_notes, parse_slide_hidden
from helpers.slide_model import normalize_element

# Chrome constants mirrored from template/lib/design.py (in inches).
def _map_color(value):
    """Map a parsed fill/line color to a Python literal."""
    if value is None:
        return None
    if value == "none":
        return "none"
    if isinstance(value, dict):
        # Gradient/pattern dicts keep their structure; only their inner colors
        # are mapped. Color+alpha dicts are preserved as-is.
        if value.get("type") == "gradient":
            return {
                **value,
                "stops": [
                    {k: _map_color(v) if k == "color" else v for k, v in stop.items()}
                    for stop in value.get("stops", [])
                ],
            }
        if value.get("type") == "pattern":
            return {
                **value,
                "fg": _map_color(value.get("fg")),
                "bg": _map_color(value.get("bg")),
            }
        inner = _map_color(value.get("color"))
        return {"color": inner, "alpha": value.get("alpha")}
    if isinstance(value, str) and value.startswith("scheme:"):
        return f"theme_{value.split(':', 1)[1]}"
    return value


def _style_color(shape, ref_key):
    """Map a <p:style> reference (fillRef/lnRef/effectRef) to a theme string."""
    style = shape.get("style") or {}
    ref = style.get(ref_key)
    if ref and ref.get("color"):
        return _map_color(ref["color"])
    return None


def _align_literal(algn):
    mapping = {
        "l": "PP_ALIGN.LEFT",
        "r": "PP_ALIGN.RIGHT",
        "ctr": "PP_ALIGN.CENTER",
        "just": "PP_ALIGN.JUSTIFY",
        "dist": "PP_ALIGN.DISTRIBUTE",
    }
    return mapping.get(algn)


def _anchor_literal(anchor):
    mapping = {
        "t": "MSO_ANCHOR.TOP",
        "ctr": "MSO_ANCHOR.MIDDLE",
        "b": "MSO_ANCHOR.BOTTOM",
    }
    return mapping.get(anchor)


def _font_size(points_hundredths):
    if points_hundredths is None:
        return None
    try:
        return round(int(points_hundredths) / 100)
    except (ValueError, TypeError):
        return None


def _shape_kind(geom_type):
    """Map a:prstGeom/@prst to the string kind passed to shapes.add_shape."""
    if geom_type in ("rect", "roundRect", "line"):
        return geom_type
    mapping = {
        "oval": "oval",
        "ellipse": "oval",
        "triangle": "triangle",
        "rightTriangle": "rightTriangle",
        "parallelogram": "parallelogram",
        "trapezoid": "trapezoid",
        "chevron": "chevron",
        "pentagon": "pentagon",
        "hexagon": "hexagon",
        "octagon": "octagon",
        "diamond": "diamond",
        "cross": "cross",
        "heart": "heart",
        "star5": "star5",
        "star6": "star6",
        "star8": "star8",
        "star10": "star10",
        "star12": "star12",
        "rightArrow": "rightArrow",
        "leftArrow": "leftArrow",
        "upArrow": "upArrow",
        "downArrow": "downArrow",
        "leftRightArrow": "leftRightArrow",
        "upDownArrow": "upDownArrow",
        "bentArrow": "bentArrow",
        "curvedRightArrow": "curvedRightArrow",
        "curvedLeftArrow": "curvedLeftArrow",
        "rightBrace": "rightBrace",
        "leftBrace": "leftBrace",
        "sun": "sun",
        "cloud": "cloud",
        "smileyFace": "smileyFace",
        "noSymbol": "noSymbol",
        "can": "can",
        "cube": "cube",
        "bevel": "bevel",
        "foldedCorner": "foldedCorner",
        "frame": "frame",
        "plaque": "plaque",
        "donut": "donut",
        "arc": "arc",
        "blockArc": "blockArc",
        "chord": "chord",
        "pie": "pie",
        "teardrop": "teardrop",
        "wave": "wave",
        "doubleWave": "doubleWave",
    }
    return mapping.get(geom_type)


def _adjustments(geom):
    """Return adjustment list for a geometry dict."""
    if not geom or "adj" not in geom:
        return None
    return [(name, fmla) for name, fmla in geom["adj"]]


def _rotation_deg(shape):
    """Return a shape's rotation in degrees, or None if unrotated/invalid."""
    rot = shape.get("rot")
    if not rot or str(rot) == "0":
        return None
    try:
        deg = round(int(rot) / 60000, 3)
    except (TypeError, ValueError):
        return None
    return deg or None


# Preset geometry that python-pptx emits for each connector ``kind``. When the
# source uses a different variant (e.g. ``bentConnector2`` vs ``bentConnector3``)
# we pass the exact preset through so it round-trips faithfully.
_CONNECTOR_DEFAULT_PRESET = {
    "straight": "straightConnector1",
    "elbow": "bentConnector3",
    "curved": "curvedConnector3",
}


def _format_kwargs(kwargs):
    """Render keyword arguments for a helper call."""
    parts = []
    for key, value in kwargs.items():
        if value is None:
            parts.append(f"{key}=None")
        elif isinstance(value, bool):
            parts.append(f"{key}={value}")
        elif isinstance(value, str) and (
            value.startswith("PP_ALIGN.") or value.startswith("MSO_ANCHOR.")
        ):
            parts.append(f"{key}={value}")
        else:
            parts.append(f"{key}={value!r}")
    return ", ".join(parts)


# ---------------------------------------------------------------------------
# Run / paragraph formatting
# ---------------------------------------------------------------------------

_RUN_ATTRS = [
    ("font", "typeface"),
    ("size", "sz"),
    ("color", "color"),
    ("bold", "b"),
    ("italic", "i"),
    ("underline", "u"),
    ("strike", "strike"),
    ("baseline", "baseline"),
    ("spacing", "spc"),
    ("highlight", "highlight"),
    ("hyperlink", "hyperlink"),
    ("effects", "effects"),
    ("pitch_family", "typeface_pitchFamily"),
    ("charset", "typeface_charset"),
]


def _fmt_run_attr(out_key, val):
    """Convert a raw parsed run attribute to a Python-friendly value."""
    if val is None or val == "":
        return None
    if out_key == "size":
        try:
            return round(int(val) / 100)
        except (ValueError, TypeError):
            return None
    if out_key in ("bold", "italic"):
        return val == "1" or val is True
    if out_key == "underline":
        if val in (False, "none", "0"):
            return False
        if val == "1":
            return True
        return val
    if out_key == "strike":
        if val in (False, "none", "0"):
            return False
        if val == "1":
            return "sngStrike"
        return val
    if out_key == "baseline":
        try:
            return int(val)
        except (ValueError, TypeError):
            return None
    if out_key == "spacing":
        try:
            return int(val)
        except (ValueError, TypeError):
            return None
    if out_key == "color":
        return _map_color(val)
    if out_key == "highlight":
        return _map_color(val)
    return val


def _run_spec_dict(run):
    """Return a run spec dict, always including font (None if absent)."""
    if "math_xml" in run:
        return {"math_xml": run["math_xml"]}
    spec = {"text": run.get("text", "")}
    for out_key, src_key in _RUN_ATTRS:
        val = _fmt_run_attr(out_key, run.get(src_key))
        if val is not None:
            spec[out_key] = val
    return spec


def _run_formatting(run):
    """Return formatting attributes for a run."""
    spec = {}
    for out_key, src_key in _RUN_ATTRS:
        val = _fmt_run_attr(out_key, run.get(src_key))
        if val is not None:
            spec[out_key] = val
    return spec


def _paragraph_props(p):
    """Return paragraph-level properties."""
    props = {}
    if p.get("algn") is not None:
        props["align"] = _align_literal(p["algn"]) or p["algn"]
    if p.get("marL") is not None:
        props["marL"] = int(p["marL"])
    if p.get("indent") is not None:
        props["indent"] = int(p["indent"])
    if p.get("bullet") is not None:
        bullet = p["bullet"]
        if bullet == "char" and p.get("bullet_char"):
            props["bullet"] = {"type": "char", "char": p["bullet_char"]}
        elif bullet == "autoNum" and p.get("bullet_type"):
            props["bullet"] = {"type": "autoNum", "style": p["bullet_type"]}
        elif bullet == "blip":
            # Picture bullets can't be emitted from python-pptx; use a dot.
            props["bullet"] = {"type": "char", "char": "•"}
        else:
            props["bullet"] = bullet
    if p.get("lnSpc") is not None:
        ln = p["lnSpc"]
        if isinstance(ln, str) and ln.endswith("pts"):
            props["line_spacing"] = ln
        else:
            props["line_spacing"] = int(ln)
    # spaceBefore/spaceAfter: only point values are emitted. Percentage spacing
    # is skipped because python-pptx's paragraph spacing API works in points.
    for src_key, dst_key in [("spaceBefore", "space_before"), ("spaceAfter", "space_after")]:
        v = p.get(src_key)
        if v is None:
            continue
        if isinstance(v, str) and v.endswith("pts"):
            props[dst_key] = int(v.replace("pts", "").strip()) / 100
    return props


def _common_paragraph_props(paragraphs):
    """Return paragraph properties that are identical across all paragraphs."""
    if not paragraphs:
        return {}
    props_list = [_paragraph_props(p) for p in paragraphs]
    first = dict(props_list[0])
    for props in props_list[1:]:
        for key in list(first.keys()):
            if first.get(key) != props.get(key):
                first.pop(key, None)
    # Alignment is stored on the paragraph object, not in _paragraph_props.
    alignments = [p.get("algn") for p in paragraphs]
    if len(set(alignments)) == 1 and alignments[0] is not None:
        first["align"] = _align_literal(alignments[0]) or alignments[0]
    return first


def _common_run_formatting(runs):
    """Return formatting attrs shared by every text run in the paragraph."""
    if not runs:
        return {}
    text_runs = [r for r in runs if "math_xml" not in r]
    if not text_runs:
        return {}
    specs = [_run_spec_dict(r) for r in text_runs]
    keys = [out_key for out_key, _ in _RUN_ATTRS]
    common = {}
    first = specs[0]
    for key in keys:
        val = first.get(key)
        if val is None:
            continue
        if all(spec.get(key) == val for spec in specs):
            common[key] = val
    return common


def _paragraphs_arg(paragraphs):
    """Return a Python literal for a list of paragraph dicts.

    Run formatting is always attached to individual runs so mixed paragraphs
    round-trip correctly. Formatting shared by every run in a paragraph is
    promoted to paragraph-level defaults to keep the generated code concise.
    Paragraph-level properties (alignment, bullets, spacing, indentation) are
    attached to the paragraph dict.
    """
    if not paragraphs:
        return ""

    out = []
    for p in paragraphs:
        para_props = _paragraph_props(p)
        runs = [
            r for r in p.get("runs", [])
            if r.get("math_xml") is not None or r.get("text") not in (None, "", "\n")
        ]

        # Empty paragraph -> empty string.
        if not runs:
            if para_props:
                out.append({"text": "", **para_props})
            else:
                out.append("")
            continue

        common_fmt = _common_run_formatting(runs)

        run_specs = []
        for r in runs:
            if "math_xml" in r:
                run_specs.append({"math_xml": r["math_xml"]})
                continue
            spec = _run_spec_dict(r)
            for key in common_fmt:
                spec.pop(key, None)
            if len(spec) == 1:
                run_specs.append(spec["text"])
            else:
                run_specs.append(spec)

        # Single unformatted text run -> plain string for concise code.
        if (
            len(run_specs) == 1
            and isinstance(run_specs[0], str)
            and not common_fmt
            and not para_props
        ):
            out.append(run_specs[0])
            continue

        para_dict = {"text": run_specs}
        para_dict.update(common_fmt)
        para_dict.update(para_props)
        out.append(para_dict)

    if len(out) == 1 and isinstance(out[0], str):
        return out[0]
    return out


# ---------------------------------------------------------------------------
# Background / notes
# ---------------------------------------------------------------------------

def _code_for_background(bg):
    if bg is None:
        return None
    if bg.get("type") == "solid":
        color = _map_color(bg.get("color"))
        if color:
            return f"shapes.add_background(slide, {color!r})"
    if bg.get("type") == "gradient":
        stops = bg.get("stops", [])
        if stops:
            parsed = [
                (stop.get("pos", i / max(len(stops) - 1, 1)), _map_color(stop.get("color")) or stop.get("color"))
                for i, stop in enumerate(stops)
            ]
            angle = bg.get("angle", 0)
            return f"shapes.add_background(slide, {{'type': 'gradient', 'angle': {angle!r}, 'stops': {parsed!r}}})"

    return None


def _code_for_notes(notes):
    if not notes:
        return None
    return f"shapes.add_notes(slide, {notes!r})"


def _code_for_hidden(hidden):
    if not hidden:
        return None
    return "shapes.set_slide_hidden(slide)"


# ---------------------------------------------------------------------------
# Shape code generators
# ---------------------------------------------------------------------------

def _common_text_kwargs(shape):
    """Return kwargs shared by add_box/add_text/add_label."""
    kwargs = {}
    paragraphs = shape.get("paragraphs", [])
    text_arg = _paragraphs_arg(paragraphs)
    if text_arg != "":
        kwargs["text"] = text_arg

    # Only use paragraph properties shared by every paragraph as top-level
    # defaults; mixed properties stay in the per-paragraph dicts.
    common_para_props = _common_paragraph_props(paragraphs)
    if common_para_props.get("align"):
        align_lit = _align_literal(common_para_props["align"])
        if align_lit:
            kwargs["align"] = align_lit
    if common_para_props.get("lnSpc") is not None:
        ln = common_para_props["lnSpc"]
        kwargs["line_spacing"] = ln if isinstance(ln, str) and ln.endswith("pts") else int(ln)
    if common_para_props.get("spaceBefore") is not None:
        kwargs["space_before"] = int(common_para_props["spaceBefore"]) / 100
    if common_para_props.get("spaceAfter") is not None:
        kwargs["space_after"] = int(common_para_props["spaceAfter"]) / 100
    if common_para_props.get("bullet") is not None:
        kwargs["bullet"] = common_para_props["bullet"]

    anchor_lit = _anchor_literal(shape.get("anchor"))
    if anchor_lit:
        kwargs["anchor"] = anchor_lit

    if shape.get("margins"):
        kwargs["margins"] = shape["margins"]
    if shape.get("wrap") is not None:
        kwargs["wrap"] = shape["wrap"] == "square"
    if shape.get("autofit") is not None:
        kwargs["autofit"] = shape["autofit"]

    return kwargs


def _line_kwargs(shape):
    """Return line-related kwargs from a parsed line dict."""
    line = shape.get("line") or {}
    kwargs = {}
    if line.get("color") is not None:
        kwargs["line"] = _map_color(line.get("color"))
    if line.get("w") is not None:
        kwargs["line_width"] = round(line["w"], 2)
    if line.get("dash") is not None and line.get("dash") != "solid":
        kwargs["line_dash"] = line["dash"]
    if line.get("head") is not None:
        kwargs["line_head"] = line["head"]
    if line.get("tail") is not None:
        kwargs["line_tail"] = line["tail"]
    if line.get("cap") is not None:
        kwargs["line_cap"] = line["cap"]
    if line.get("cmpd") is not None:
        kwargs["line_cmpd"] = line["cmpd"]
    return kwargs


def _style_and_geom_kwargs(shape):
    """Return style, adjustments, rotation, flip, effects kwargs."""
    kwargs = {}
    geom = shape.get("geom", {})
    adj = _adjustments(geom)
    if adj:
        kwargs["adjustments"] = adj
    if shape.get("style"):
        kwargs["style"] = shape["style"]
    if shape.get("rot") and shape.get("rot") != "0":
        kwargs["rotation"] = int(shape["rot"]) / 60000
    if shape.get("flipH"):
        kwargs["flip_h"] = True
    if shape.get("flipV"):
        kwargs["flip_v"] = True
    if shape.get("effects"):
        kwargs["effects"] = shape["effects"]
    return kwargs


def _code_for_line_or_arrow(shape, x, y, w, h):
    line = shape.get("line") or {}
    color = _map_color(line.get("color")) or d_col("blue")
    width = line.get("w", 1.0)
    flip_h = shape.get("flipH") == "1"
    flip_v = shape.get("flipV") == "1"

    x1 = x + w if flip_h else x
    y1 = y + h if flip_v else y
    x2 = x if flip_h else x + w
    y2 = y if flip_v else y + h

    kwargs = {"color": color, "width": width}
    if line.get("dash") and line["dash"] != "solid":
        kwargs["dash"] = line["dash"]
    if line.get("head"):
        kwargs["head"] = line["head"]
    if line.get("tail"):
        kwargs["tail"] = line["tail"]
    if line.get("cap"):
        kwargs["cap"] = line["cap"]
    if line.get("cmpd"):
        kwargs["cmpd"] = line["cmpd"]
    if shape.get("style"):
        kwargs["style"] = shape["style"]
    rotation = _rotation_deg(shape)
    if rotation is not None:
        kwargs["rotation"] = rotation

    return (
        f"shapes.add_line(slide, {x1:.3f}, {y1:.3f}, {x2:.3f}, {y2:.3f}, "
        f"{_format_kwargs(kwargs)})"
    )


def _connector_kind(geom_type):
    """Map a connector preset geometry to the kind string used by add_connector."""
    if not geom_type:
        return "straight"
    if "bent" in geom_type or "elbow" in geom_type:
        return "elbow"
    if "curve" in geom_type or "curved" in geom_type:
        return "curved"
    return "straight"


def _code_for_connector(shape, x, y, w, h):
    line = shape.get("line") or {}
    color = _map_color(line.get("color")) or d_col("blue")
    width = line.get("w", 1.0)
    flip_h = shape.get("flipH") == "1"
    flip_v = shape.get("flipV") == "1"

    x1 = x + w if flip_h else x
    y1 = y + h if flip_v else y
    x2 = x if flip_h else x + w
    y2 = y if flip_v else y + h

    kwargs = {"color": color, "width": width}
    if line.get("dash") and line["dash"] != "solid":
        kwargs["dash"] = line["dash"]
    if line.get("head"):
        kwargs["head"] = line["head"]
    if line.get("tail"):
        kwargs["tail"] = line["tail"]
    if line.get("cap"):
        kwargs["cap"] = line["cap"]
    if line.get("cmpd"):
        kwargs["cmpd"] = line["cmpd"]
    if shape.get("style"):
        kwargs["style"] = shape["style"]
    geom = shape.get("geom", {})
    geom_type = geom.get("type")
    kind = _connector_kind(geom_type)
    if kind != "straight":
        kwargs["kind"] = kind
    # Preserve the exact preset (and its guides) when it differs from the
    # variant that ``kind`` maps to — otherwise an L-bend becomes a Z-bend, etc.
    if geom_type and geom_type != _CONNECTOR_DEFAULT_PRESET.get(kind):
        kwargs["preset"] = geom_type
        adj = _adjustments(geom)
        if adj:
            kwargs["adjustments"] = adj
    rotation = _rotation_deg(shape)
    if rotation is not None:
        kwargs["rotation"] = rotation

    return (
        f"shapes.add_connector(slide, {x1:.3f}, {y1:.3f}, {x2:.3f}, {y2:.3f}, "
        f"{_format_kwargs(kwargs)})"
    )


def _code_for_box(shape, x, y, w, h):
    """Emit code for add_box or add_shape."""
    geom = shape.get("geom", {})
    geom_type = geom.get("type")
    kind = _shape_kind(geom_type)

    kwargs = {}
    fill = _map_color(shape.get("fill")) or _style_color(shape, "fillRef")
    kwargs["fill"] = fill
    kwargs.update(_line_kwargs(shape))
    kwargs.update(_style_and_geom_kwargs(shape))

    # Text formatting
    text_kwargs = _common_text_kwargs(shape)
    # Avoid duplicating fill/line color as text color unless explicitly set.
    kwargs.update(text_kwargs)

    if geom_type in ("rect", "roundRect") or kind in ("rect", "roundRect"):
        rounded = geom_type == "roundRect" or kind == "roundRect"
        if not rounded:
            kwargs["rounded"] = False
        if kwargs.get("adjustments"):
            # roundRect only has one adj; extract it.
            for adj_name, fmla in kwargs["adjustments"]:
                if adj_name == "adj" and fmla.startswith("val "):
                    try:
                        kwargs["rounded_adj"] = int(fmla.split()[1])
                    except (ValueError, IndexError):
                        pass
            kwargs.pop("adjustments", None)
        text = kwargs.pop("text", "")
        chrome_name = shape.get("name")
        slide_number_expr = ""
        name_arg = ""
        if chrome_name == "SlideNumber":
            slide_number_expr = ", slide_number=n"
        elif chrome_name in ("Footer", "Date", "Header", "Title"):
            name_arg = f", name={chrome_name!r}"
        extra = ""
        if kwargs:
            extra = f", {_format_kwargs(kwargs)}"
        extra += name_arg + slide_number_expr
        return f"shapes.add_box(slide, {text!r}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f}{extra})"

    # Generic shape
    kwargs["kind"] = kind or geom_type
    text = kwargs.pop("text", "")
    return f"shapes.add_shape(slide, {kwargs['kind']!r}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f}" + (
        f", text={text!r}, {_format_kwargs({k:v for k,v in kwargs.items() if k != 'kind'})})" if kwargs else ")"
    )


def _code_for_text_shape(shape, x, y, w, h):
    """Emit code for add_label or add_text."""
    kwargs = _common_text_kwargs(shape)
    fill = _map_color(shape.get("fill"))
    line = shape.get("line") or {}
    line_color = _map_color(line.get("color"))
    helper = "add_label"
    if fill is not None or line_color is not None:
        helper = "add_text"
        kwargs["fill"] = fill
        kwargs.update(_line_kwargs(shape))

    text = kwargs.pop("text", "")
    chrome_name = shape.get("name")
    slide_number_expr = ""
    name_arg = ""
    if chrome_name == "SlideNumber":
        slide_number_expr = ", slide_number=n"
    elif chrome_name in ("Footer", "Date", "Header", "Title"):
        name_arg = f", name={chrome_name!r}"
    extra = ""
    if kwargs:
        extra = f", {_format_kwargs(kwargs)}"
    extra += name_arg + slide_number_expr
    return f"shapes.{helper}(slide, {text!r}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f}{extra})"


def _code_for_table(shape):
    x, y, w, h = shape["x"], shape["y"], shape["w"], shape["h"]
    kwargs = {"x": x, "y": y, "w": w, "h": h, "borders": False}
    if shape.get("colWidths"):
        kwargs["col_widths"] = shape["colWidths"]
    if shape.get("rowHeights"):
        kwargs["row_heights"] = shape["rowHeights"]

    cells = []
    for ri, row in enumerate(shape.get("cells", [])):
        cell_row = []
        for ci, cell in enumerate(row):
            spec = {"text": cell.get("text", "")}
            if cell.get("paragraphs"):
                spec["text"] = _paragraphs_arg(cell["paragraphs"])
            if cell.get("fill") is not None:
                spec["fill"] = _map_color(cell["fill"])
            if cell.get("anchor"):
                spec["anchor"] = _anchor_literal(cell["anchor"])
            if cell.get("margins"):
                spec["margins"] = cell["margins"]
            if cell.get("borders"):
                spec["borders"] = {
                    side: {
                        "color": _map_color(b.get("color")) or "000000",
                        "w": b.get("w", 0.5),
                        "dash": b.get("dash", "solid"),
                    }
                    for side, b in cell["borders"].items()
                }
            if cell.get("gridSpan"):
                spec["gridSpan"] = cell["gridSpan"]
            if cell.get("rowSpan"):
                spec["rowSpan"] = cell["rowSpan"]
            cell_row.append(spec)
        cells.append(cell_row)

    kwargs["cells"] = cells
    return f"shapes.add_custom_table(slide, {_format_kwargs(kwargs)})"


def _code_for_image(shape, media_names):
    raw = shape.get("imgFile", "")
    name = media_names.get(raw) if media_names else None
    if name is None:
        name = Path(raw).name if raw else raw
    if name and not Path(name).suffix and Path(raw).suffix:
        name = name + Path(raw).suffix
    if not name:
        shape_name = shape.get("name", "Picture")
        return f"# TODO: image asset not found for '{shape_name}' - verify asset extraction"
    x, y, w, h = shape["x"], shape["y"], shape["w"], shape["h"]
    kwargs = {}
    img_info = shape.get("imgInfo") or {}
    crop = img_info.get("crop")
    if crop:
        kwargs["crop"] = crop
    lum = img_info.get("lum")
    if lum:
        kwargs["lum"] = lum
    return f"shapes.add_image(slide, {name!r}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f}" + (
        f", {_format_kwargs(kwargs)})" if kwargs else ")"
    )


def _code_for_movie(shape, media_names):
    raw = shape.get("mediaFile", "")
    name = media_names.get(raw) if media_names else None
    if name is None:
        name = Path(raw).name if raw else raw
    if name and not Path(name).suffix and Path(raw).suffix:
        name = name + Path(raw).suffix
    if not name:
        shape_name = shape.get("name", "Movie")
        return f"# TODO: media asset not found for '{shape_name}' - verify asset extraction"
    poster_raw = shape.get("posterFile", "")
    poster_name = media_names.get(poster_raw) if media_names else None
    if poster_name is None:
        poster_name = Path(poster_raw).name if poster_raw else None
    if poster_name and not Path(poster_name).suffix and Path(poster_raw).suffix:
        poster_name = poster_name + Path(poster_raw).suffix
    x, y, w, h = shape["x"], shape["y"], shape["w"], shape["h"]
    kwargs = {"poster_name": poster_name} if poster_name else {}
    return f"shapes.add_movie(slide, {name!r}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f}" + (
        f", {_format_kwargs(kwargs)})" if kwargs else ")"
    )


def _code_for_chart(shape):
    x, y, w, h = shape["x"], shape["y"], shape["w"], shape["h"]
    chart_type = shape.get("chart_type") or "COLUMN_CLUSTERED"
    categories = shape.get("categories", [])
    series = shape.get("series", [])
    title = shape.get("title", "")
    kwargs = {"title": title} if title else {}
    return (
        f"shapes.add_chart(slide, {chart_type!r}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f}, "
        f"{categories!r}, {series!r}, {_format_kwargs(kwargs)})"
    )


def _code_for_group(shape, media_names, target_var="slide", group_var="grp", assets_dir=None, capture=None):
    x, y, w, h = shape["x"], shape["y"], shape["w"], shape["h"]
    lines = [f"{group_var} = shapes.add_group({target_var}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f})"]
    child_group_var = group_var + "_"
    for child in shape.get("children", []):
        if child.get("type") == "group":
            # Bind a captured child group to its connection variable.
            child_var = (capture or {}).get(str(child.get("id")))
            code = _code_for_group(child, media_names, target_var=group_var, group_var=child_var or child_group_var, assets_dir=assets_dir, capture=capture)
            child_group_var += "_"
        else:
            code = _code_for_any(child, media_names, group_var=group_var, assets_dir=assets_dir, capture=capture)
        if code:
            lines.append(code)
    lines.append(f"shapes.set_group_bounds({group_var}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f})")
    return "\n".join(lines)


def _apply_flip_to_svg(svg_data: str, flip_h: bool, flip_v: bool) -> str:
    """Wrap an SVG's content in a transform that mirrors it horizontally
    and/or vertically around the viewBox center."""
    if not flip_h and not flip_v:
        return svg_data
    m = re.search(r'viewBox="0 0 ([\d.]+) ([\d.]+)"', svg_data)
    if not m:
        return svg_data
    vw, vh = float(m.group(1)), float(m.group(2))
    if flip_h and flip_v:
        transform = f"translate({vw}, {vh}) scale(-1, -1)"
    elif flip_h:
        transform = f"translate({vw}, 0) scale(-1, 1)"
    else:
        transform = f"translate(0, {vh}) scale(1, -1)"
    # Insert a wrapping group just after the opening <svg ...> tag.
    head_match = re.search(r'(<svg[^>]*>)', svg_data)
    if not head_match:
        return svg_data
    head = head_match.group(1)
    tail = '</svg>'
    inner = svg_data[head_match.end():svg_data.rfind(tail)]
    return f'{head}<g transform="{transform}">{inner}</g>{tail}'


def _code_for_freeform_svg(shape, x, y, w, h, assets_dir, group_var="slide"):
    """Write a custom-geometry shape's SVG to assets/ and emit add_image."""
    svg_data = shape.get("geom", {}).get("svg_data")
    if not svg_data or not assets_dir:
        return _unsupported_comment(shape)
    assets_dir = Path(assets_dir)
    assets_dir.mkdir(parents=True, exist_ok=True)

    flip_h = shape.get("flipH") in (True, "1", 1)
    flip_v = shape.get("flipV") in (True, "1", 1)
    if flip_h or flip_v:
        svg_data = _apply_flip_to_svg(svg_data, flip_h, flip_v)

    name = f"freeform_{hashlib.sha1(svg_data.encode('utf-8')).hexdigest()[:12]}.svg"
    asset_path = assets_dir / name
    if not asset_path.exists():
        asset_path.write_text(svg_data, encoding="utf-8")
    return f"shapes.add_image({group_var}, {name!r}, {x:.3f}, {y:.3f}, {w:.3f}, {h:.3f})"


def _code_for_any(shape, media_names=None, group_var="slide", assets_dir=None, capture=None):
    shape_type = shape.get("type")
    x, y, w, h = shape["x"], shape["y"], shape["w"], shape["h"]
    geom = shape.get("geom", {})
    geom_type = geom.get("type")
    # Variable to bind this shape to, when it is a connector or a connection
    # target that a `connect_shapes` call later references.
    my_var = (capture or {}).get(str(shape.get("id")))

    # Custom geometry that we successfully converted to SVG -> render as image.
    if geom_type == "custom" and geom.get("svg_data"):
        code = _code_for_freeform_svg(shape, x, y, w, h, assets_dir, group_var=group_var)
        return f"{my_var} = {code}" if my_var else code

    if shape_type == "image":
        code = _code_for_image(shape, media_names)
    elif shape_type == "movie":
        code = _code_for_movie(shape, media_names)
    elif shape_type == "chart":
        code = _code_for_chart(shape)
    elif shape_type in ("line", "arrow") or (shape_type == "shape" and geom_type == "line"):
        code = _code_for_line_or_arrow(shape, x, y, w, h)
    elif shape_type == "connector":
        code = _code_for_connector(shape, x, y, w, h)
    elif shape_type == "table":
        code = _code_for_table(shape)
    elif shape_type == "textbox":
        code = _code_for_text_shape(shape, x, y, w, h)
    elif shape_type == "group":
        # A captured group binds its own variable inside _code_for_group; return
        # directly so the assignment isn't double-wrapped.
        gv = my_var or "grp"
        return _code_for_group(
            shape, media_names, target_var=group_var, group_var=gv,
            assets_dir=assets_dir, capture=capture,
        )
    elif shape_type in ("shape", "auto_shape", "placeholder") or "shape" in str(shape_type).lower():
        # Property-first classification: placeholders are treated as normal shapes.
        eff_fill = shape.get("fill")
        eff_line = shape.get("line")
        has_text = bool(shape.get("text"))
        has_known_geom = geom_type in ("rect", "roundRect") or _shape_kind(geom_type) is not None

        if has_text and (eff_fill is not None or eff_line is not None or has_known_geom):
            # Text with a visible container or known geometry -> render as a box.
            code = _code_for_box(shape, x, y, w, h)
        elif has_text:
            # Plain text placeholder / text box -> transparent text shape.
            code = _code_for_text_shape(shape, x, y, w, h)
        elif eff_fill is not None or eff_line is not None or has_known_geom:
            # Visual shape without text -> render as a box/shape.
            code = _code_for_box(shape, x, y, w, h)
        else:
            code = _unsupported_comment(shape)
    else:
        code = _unsupported_comment(shape)

    if group_var != "slide":
        # Redirect slide-level helper calls to the group variable.
        code = code.replace("(slide,", f"({group_var},")
        code = code.replace(", slide,", f", {group_var},")
    if my_var and code and not code.lstrip().startswith("#"):
        code = f"{my_var} = {code}"
    return code


def _unsupported_comment(shape):
    name = shape.get("name", "shape")
    return f"    # TODO: unsupported {shape['type']} '{name}' - implement manually"


def d_col(key):
    """Fallback design palette lookup for defaults not extractable from shape."""
    return {
        "blue": "4F81BD",
        "green": "9BBB59",
        "sep": "BFBFBF",
    }.get(key, key)


# ---------------------------------------------------------------------------
# Chrome detection
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def _shape_text(shape):
    """Return the stripped plain text of a normalized shape."""
    paragraphs = shape.get("paragraphs") or []
    parts = []
    for para in paragraphs:
        if isinstance(para, str):
            parts.append(para)
            continue
        runs = para.get("runs") if isinstance(para.get("runs"), list) else (
            para.get("text") if isinstance(para.get("text"), list) else []
        )
        for r in runs:
            if isinstance(r, dict):
                parts.append(r.get("text", ""))
            else:
                parts.append(str(r))
    return "".join(parts).strip()


def _set_shape_text(shape, new_text):
    """Replace a shape's paragraph text while preserving run formatting."""
    paragraphs = shape.get("paragraphs") or []
    if not paragraphs:
        shape["paragraphs"] = [{"text": [{"text": new_text}]}]
        return
    first = paragraphs[0]
    runs = first.get("runs") if isinstance(first.get("runs"), list) else (
        first.get("text") if isinstance(first.get("text"), list) else []
    )
    if runs and isinstance(runs[0], dict):
        runs[0]["text"] = new_text
        for r in runs[1:]:
            r["text"] = ""
    elif runs:
        first["text"] = [{"text": new_text}]
    else:
        first["text"] = [{"text": new_text}]
    shape["paragraphs"] = [first]


def _is_slide_number_text(text):
    return text.strip() == "\u2039#\u203a"  # PowerPoint slide-number marker ‹#›


def detect_footer_text(layout_xmls: list[Path]) -> str | None:
    """Infer the footer text by finding a short bottom-of-slide string repeated across layouts."""
    counts = {}
    for xml_path in layout_xmls:
        for raw in read_slide_shapes(xml_path):
            s = normalize_element(raw)
            if s.get("placeholder"):
                continue
            # Footer text lives near the bottom of the layout.
            if s.get("y", 0) < 6.0:
                continue
            text = _shape_text(s)
            if not text or _is_slide_number_text(text):
                continue
            counts[text] = counts.get(text, 0) + 1
    # Require the text to appear in at least two layouts and be reasonably short.
    candidates = [(t, c) for t, c in counts.items() if c >= 2 and len(t) <= 120]
    if not candidates:
        return None
    # Most frequent, then shortest.
    return sorted(candidates, key=lambda x: (-x[1], len(x[0])))[0][0]


def generate_layout_chrome_code(layout_xml: Path, media_names: dict, slide_num: int, footer_text: str | None = None, assets_dir: Path | None = None) -> str:
    """Generate code for layout-level footer/page-number shapes on a slide.

    These shapes live on the slide layout, so PowerPoint locks them on the
    slide. By emitting them as normal slide shapes we make them selectable
    and keep them in sync when a slide is regenerated.
    """
    raw_shapes = read_slide_shapes(layout_xml)
    shapes = [normalize_element(s) for s in raw_shapes]
    shapes.sort(key=lambda s: s.get("z", 0))

    has_regular_footer = footer_text and any(
        not s.get("placeholder") and _shape_text(s) == footer_text
        for s in shapes
    )
    has_regular_sldnum = any(
        not s.get("placeholder") and _is_slide_number_text(_shape_text(s))
        for s in shapes
    )

    lines = []
    for s in shapes:
        ph = s.get("placeholder")
        text = _shape_text(s)
        # Footer / slide-number chrome lives near the bottom of the layout.
        y = s.get("y", 0)
        included = False
        if ph:
            if ph == "sldNum" and not has_regular_sldnum and y >= 6.0:
                _set_shape_text(s, str(slide_num))
                s["name"] = "SlideNumber"
                included = True
            elif ph in ("ftr", "dt", "hdr") and footer_text and not has_regular_footer and y >= 6.0:
                _set_shape_text(s, footer_text)
                s["name"] = {"ftr": "Footer", "dt": "Date", "hdr": "Header"}.get(ph)
                included = True
        else:
            if _is_slide_number_text(text) and y >= 6.0:
                _set_shape_text(s, str(slide_num))
                s["name"] = "SlideNumber"
                included = True
            elif footer_text and text == footer_text and y >= 6.0:
                s["name"] = "Footer"
                included = True

        if not included:
            continue

        code = _code_for_any(s, media_names, assets_dir=assets_dir)
        if s.get("name") == "Footer" and footer_text:
            # Reference the central design constant instead of a baked-in literal,
            # so editing design.FOOTER_TEXT updates the footer on every slide.
            code = code.replace(repr(footer_text), "d.FOOTER_TEXT", 1)
        if code and not code.lstrip().startswith("# TODO"):
            lines.append(code)

    if not lines:
        return ""

    out = ["    # Layout chrome (footer / slide number)"]
    for code in lines:
        for line in code.split("\n"):
            out.append(f"    {line}")
    return "\n".join(out)


def _sanitize_id(value) -> str:
    """Make a shape id safe to use as a Python identifier suffix."""
    return re.sub(r"\W", "_", str(value))


def _collect_shape_ids(shapes, out):
    """Collect the ids of every shape in the tree (recursing into groups)."""
    for s in shapes:
        if s.get("id"):
            out.add(str(s["id"]))
        if s.get("children"):
            _collect_shape_ids(s["children"], out)


def _collect_connectors(shapes, out):
    """Collect connectors that carry shape-to-shape connections."""
    for s in shapes:
        if s.get("type") == "connector" and s.get("connections"):
            out.append(s)
        if s.get("children"):
            _collect_connectors(s["children"], out)


def _connection_plan(shapes):
    """Plan variable capture + ``connect_shapes`` calls for connected connectors.

    Returns ``(capture, connect_lines)``: ``capture`` maps a source shape id to
    the Python variable the generated code should bind it to, and
    ``connect_lines`` are the trailing ``shapes.connect_shapes(...)`` statements
    that re-attach each connector once every shape exists.
    """
    present = set()
    _collect_shape_ids(shapes, present)
    connectors = []
    _collect_connectors(shapes, connectors)

    capture = {}
    connect_lines = []
    for conn in connectors:
        cid = conn.get("id")
        if not cid:
            continue
        sides = []
        for side in ("begin", "end"):
            spec = conn["connections"].get(side)
            if not spec:
                continue
            tid = str(spec.get("id"))
            if tid not in present:
                continue  # target isn't reproduced (e.g. a placeholder) -> skip
            capture[tid] = f"_sh{_sanitize_id(tid)}"
            try:
                idx = int(spec.get("idx", 0))
            except (TypeError, ValueError):
                idx = 0
            sides.append(f"{side}=({capture[tid]}, {idx})")
        if not sides:
            continue
        cvar = f"_cx{_sanitize_id(cid)}"
        capture[str(cid)] = cvar
        connect_lines.append(f"shapes.connect_shapes({cvar}, {', '.join(sides)})")
    return capture, connect_lines


def generate_slide_code(slide_xml: Path, media_names: dict, title: str, assets_dir: Path | None = None) -> str:
    """Return the body of an add_slide() function as a code string."""
    raw_shapes = read_slide_shapes(slide_xml)
    shapes = [normalize_element(s) for s in raw_shapes]
    shapes.sort(key=lambda s: s.get("z", 0))

    lines = []
    hidden_code = _code_for_hidden(parse_slide_hidden(slide_xml))
    if hidden_code:
        lines.append(hidden_code)

    bg_code = _code_for_background(parse_background(slide_xml))
    if bg_code:
        lines.append(bg_code)

    notes_code = _code_for_notes(parse_slide_notes(slide_xml))
    if notes_code:
        lines.append(notes_code)

    capture, connect_lines = _connection_plan(shapes)

    for shape in shapes:
        if shape.get("placeholder") in ("title", "ctrTitle"):
            shape["name"] = "Title"
        lines.append(_code_for_any(shape, media_names, assets_dir=assets_dir, capture=capture))

    # Re-attach connected connectors after every shape exists (order-independent).
    lines.extend(connect_lines)

    flattened = []
    for item in lines:
        flattened.extend(str(item).split("\n"))
    return "\n".join(f"    {line}" for line in flattened)
