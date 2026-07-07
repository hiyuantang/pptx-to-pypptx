"""Canonical slide model consumed by the slide-code generator.

This module takes the raw dictionaries produced by `slide_xml.py` and normalizes
them so that style-derived effective properties, default run properties, and
other implicit values are explicit. `slide_codegen.py` consumes this model
directly when emitting python-pptx code.
"""

from typing import Any


EMU_PER_INCH = 914400


def _scheme_to_theme(color: Any) -> Any:
    """Convert 'scheme:accent1' to 'theme_accent1', preserving alpha if present."""
    if isinstance(color, dict):
        inner = color.get("color")
        if isinstance(inner, str) and inner.startswith("scheme:"):
            return {"color": "theme_" + inner.split(":", 1)[1], "alpha": color.get("alpha")}
        return color
    if isinstance(color, str) and color.startswith("scheme:"):
        return "theme_" + color.split(":", 1)[1]
    return color


def _normalize_dict_fill(fill: dict) -> dict:
    """Convert scheme colors inside gradient/pattern dict fills to theme names."""
    out = dict(fill)
    if out.get("type") == "gradient":
        out["stops"] = [
            {k: _scheme_to_theme(v) if k == "color" else v for k, v in stop.items()}
            for stop in fill.get("stops", [])
        ]
    elif out.get("type") == "pattern":
        for key in ("fg", "bg"):
            if out.get(key) is not None:
                out[key] = _scheme_to_theme(out[key])
    return out


def _effective_fill(elem: dict) -> Any:
    """Return the effective fill for a shape, resolving p:style/fillRef."""
    raw = elem.get("fill")
    if raw is not None:
        if isinstance(raw, dict):
            out = dict(raw)
            if out.get("type") == "gradient":
                out["stops"] = [
                    {k: _scheme_to_theme(v) if k == "color" else v for k, v in stop.items()}
                    for stop in raw.get("stops", [])
                ]
            elif out.get("type") == "pattern":
                for key in ("fg", "bg"):
                    if out.get(key) is not None:
                        out[key] = _scheme_to_theme(out[key])
            return out
        return _scheme_to_theme(raw)
    style = elem.get("style") or {}
    ref = style.get("fillRef")
    if ref and ref.get("color"):
        return _scheme_to_theme(ref["color"])
    # No fill element and no style reference means transparent.
    return "none"


def _effective_line(elem: dict) -> dict | None:
    """Return the effective line dict for a shape, resolving p:style/lnRef."""
    line = elem.get("line")
    if line is not None:
        out = dict(line)
        if out.get("color") is not None:
            out["color"] = _scheme_to_theme(out["color"])
        return out
    style = elem.get("style") or {}
    ref = style.get("lnRef")
    if ref and ref.get("color"):
        return {"color": _scheme_to_theme(ref["color"])}
    return None


def _merge_run_with_defaults(default: dict, run: dict) -> dict:
    """Apply default font props to a run, letting explicit run props win."""
    merged = dict(default)
    merged["text"] = run.get("text", "")
    for key, value in run.items():
        if value is not None:
            merged[key] = value
    return merged


def normalize_run(run: dict) -> dict:
    """Return a canonical run dict with theme colors resolved."""
    if "math_xml" in run:
        return {"math_xml": run["math_xml"]}
    out = {"text": run.get("text", "")}
    for key in (
        "sz", "b", "i", "u", "strike", "baseline", "spc",
        "typeface", "ea", "cs",
        "typeface_pitchFamily", "ea_pitchFamily", "cs_pitchFamily",
        "typeface_charset", "ea_charset", "cs_charset",
        "highlight", "effects", "hyperlink",
    ):
        if run.get(key) is not None:
            out[key] = run[key]
    if run.get("color") is not None:
        out["color"] = _scheme_to_theme(run["color"])
    return out


def normalize_paragraph(p: dict, body_defaults: dict | None = None) -> dict:
    """Return a canonical paragraph dict."""
    body_defaults = body_defaults or {}
    runs = [
        normalize_run(_merge_run_with_defaults(body_defaults, r))
        for r in p.get("runs", [])
        if r.get("text") != "\n" or r.get("math_xml") is not None
    ]
    out = {
        "text": p.get("text", ""),
        "runs": runs,
    }
    for key in ("algn", "lnSpc", "indent", "marL", "bullet",
                "bullet_char", "bullet_type", "spaceBefore", "spaceAfter",
                "bullet_size_pts", "bullet_size_pct", "bullet_font"):
        if p.get(key) is not None:
            out[key] = p[key]
    if p.get("bullet_color") is not None:
        out["bullet_color"] = _scheme_to_theme(p["bullet_color"])
    return out


def normalize_text_body(elem: dict) -> dict | None:
    """Return canonical text body properties."""
    paragraphs = elem.get("paragraphs")
    if not paragraphs:
        return None
    body_defaults = {}
    # Extract default run props from the first paragraph's merged defaults if present.
    # slide_xml.py already merges defRPr into runs, so we just normalize paragraphs.
    norm_paras = [normalize_paragraph(p) for p in paragraphs]
    text = "\n".join(p.get("text", "") for p in norm_paras)
    out = {
        "text": text,
        "paragraphs": norm_paras,
    }
    if elem.get("margins"):
        out["margins"] = elem["margins"]
    if elem.get("anchor") is not None:
        out["anchor"] = elem["anchor"]
    if elem.get("wrap") is not None:
        out["wrap"] = elem["wrap"]
    if elem.get("autofit") is not None:
        out["autofit"] = elem["autofit"]
    return out


def _parse_effect_string(s: str) -> dict:
    """Convert a compact effect string from slide_xml.py into an effect dict."""
    s = s.strip()
    if s.startswith("shadow(") and s.endswith(")"):
        return {"type": "outerShdw", "color": s[7:-1]}
    if s.startswith("innerShadow(") and s.endswith(")"):
        return {"type": "innerShdw", "color": s[12:-1]}
    if s.startswith("glow(") and s.endswith(")"):
        return {"type": "glow", "color": s[5:-1]}
    if s == "reflection":
        return {"type": "reflection"}
    if s.startswith("softEdge(") and s.endswith(")"):
        return {"type": "softEdge", "rad": s[9:-1]}
    return {"type": "unknown", "value": s}


def normalize_element(elem: dict) -> dict:
    """Return a canonical dict for any slide element."""
    out = {
        "type": elem.get("type"),
        "name": elem.get("name"),
        "placeholder": elem.get("placeholder"),
        "x": elem.get("x"),
        "y": elem.get("y"),
        "w": elem.get("w"),
        "h": elem.get("h"),
        "z": elem.get("z"),
    }
    for key in ("rot", "flipH", "flipV"):
        if elem.get(key) is not None:
            out[key] = elem[key]
    # Preserve the source shape id and any connector connections so bent/elbow
    # connectors can be re-attached to their target shapes on rebuild.
    if elem.get("id"):
        out["id"] = elem["id"]
    if elem.get("connections"):
        out["connections"] = elem["connections"]

    fill = _effective_fill(elem)
    if fill is not None:
        out["fill"] = fill
    line = _effective_line(elem)
    if line is not None:
        out["line"] = line
    if elem.get("effects"):
        effects = elem["effects"]
        if effects and isinstance(effects[0], str):
            effects = [_parse_effect_string(e) for e in effects]
        # Convert scheme colors inside effect dicts to theme names.
        normalized_effects = []
        for eff in effects:
            neff = dict(eff)
            if neff.get("color") is not None:
                neff["color"] = _scheme_to_theme(neff["color"])
            normalized_effects.append(neff)
        out["effects"] = normalized_effects
    if elem.get("geom"):
        geom = dict(elem["geom"])
        # Normalise line variants to a single identifier.
        if geom.get("type") in ("line", "lineInv"):
            geom["type"] = "line"
        out["geom"] = geom
    if elem.get("style"):
        out["style"] = elem["style"]
    if elem.get("hyperlink"):
        out["hyperlink"] = elem["hyperlink"]

    text_body = normalize_text_body(elem)
    if text_body is not None:
        out.update(text_body)

    if elem.get("type") == "image":
        for key in ("imgFile", "imgHash", "imgInfo"):
            if elem.get(key) is not None:
                out[key] = elem[key]

    if elem.get("type") == "table" and elem.get("cells"):
        out["rows"] = elem["rows"]
        out["cols"] = elem["cols"]
        out["colWidths"] = elem.get("colWidths", [])
        out["rowHeights"] = elem.get("rowHeights", [])
        out["cells"] = []
        for row in elem["cells"]:
            norm_row = []
            for cell in row:
                raw_margins = cell.get("margins") or {}
                margins = {}
                for k, v in raw_margins.items():
                    if v is not None:
                        try:
                            margins[k] = round(int(v) / EMU_PER_INCH, 3)
                        except (ValueError, TypeError):
                            pass
                borders = {}
                for side, b in (cell.get("borders") or {}).items():
                    borders[side] = dict(b)
                    if borders[side].get("color") is not None:
                        borders[side]["color"] = _scheme_to_theme(borders[side]["color"])
                cell_fill = cell.get("fill")
                if isinstance(cell_fill, dict):
                    cell_fill = _normalize_dict_fill(cell_fill)
                norm_cell = {
                    "text": cell.get("text", ""),
                    "fill": _scheme_to_theme(cell_fill) if cell_fill is not None else None,
                    "anchor": cell.get("anchor"),
                    "margins": margins if margins else None,
                    "borders": borders,
                    "gridSpan": cell.get("gridSpan"),
                    "rowSpan": cell.get("rowSpan"),
                }
                if cell.get("paragraphs"):
                    norm_cell["paragraphs"] = [normalize_paragraph(p) for p in cell["paragraphs"]]
                norm_row.append(norm_cell)
            out["cells"].append(norm_row)

    if elem.get("type") == "group" and elem.get("children"):
        out["children"] = [normalize_element(c) for c in elem["children"]]

    if elem.get("type") == "chart":
        for key in ("chart_type", "categories", "series", "title"):
            if elem.get(key) is not None:
                out[key] = elem[key]

    if elem.get("type") == "movie":
        for key in ("mediaFile", "posterFile"):
            if elem.get(key) is not None:
                out[key] = elem[key]

    return out


def normalize_slide(elements: list[dict]) -> list[dict]:
    """Normalize all elements on a slide."""
    return [normalize_element(e) for e in elements]



