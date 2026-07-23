"""Shared XML parsing helpers for PPTX slide shapes.

Provides a single source of truth for reading shape information from a PPTX
slide XML. Used by extract_slide.py and generate_slides.py.
"""

import hashlib
import os
import re
import xml.etree.ElementTree as ET
from pathlib import Path

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
A14 = 'http://schemas.microsoft.com/office/drawing/2010/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

# Preserve conventional prefixes when serialising Office Math XML for injection.
ET.register_namespace('m', M)
ET.register_namespace('a', A)
ET.register_namespace('a14', A14)

EMU_PER_INCH = 914400
EMU_PER_POINT = 12700


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def child_text(parent, tag_ns, tag_local, default=''):
    if parent is None:
        return default
    el = parent.find(f'{{{tag_ns}}}{tag_local}')
    if el is not None and el.text:
        return el.text
    return default


def get_name(elem, default='shape'):
    """Extract cNvPr/@name from a shape/image/connector/group."""
    for path in [
        f'{{{P}}}cNvPr',
        f'.//{{{P}}}cNvPr',
    ]:
        cNvPr = elem.find(path)
        if cNvPr is not None:
            return cNvPr.get('name', default)
    return default


def get_id(elem):
    cNvPr = elem.find(f'.//{{{P}}}cNvPr')
    if cNvPr is not None:
        return cNvPr.get('id', '')
    return ''


# ---------------------------------------------------------------------------
# Colors, fills, lines, effects
# ---------------------------------------------------------------------------

def _parse_alpha(color_el):
    """Return alpha fraction (0-1) from an <a:alpha> child, or None."""
    alpha = color_el.find(f'{{{A}}}alpha')
    if alpha is None:
        return None
    try:
        val = int(alpha.get('val', 100000))
    except (ValueError, TypeError):
        return None
    return round(val / 100000, 4)


# Luminance / saturation / tint / shade transform children. PowerPoint expresses
# most *light* colors as a darker base color plus a lightening transform (e.g. a
# theme accent with lumMod+lumOff, or an srgb base with a tint); dropping these
# silently reverts the color to its darker base on a round trip. We preserve the
# raw OOXML ``val`` strings so they can be re-emitted verbatim on rebuild.
_COLOR_MOD_TAGS = ('lumMod', 'lumOff', 'satMod', 'satOff', 'hueMod', 'hueOff', 'shade', 'tint')


def _parse_color_mods(color_el):
    """Return a dict of luminance/saturation/tint/shade transforms on a color element."""
    mods = {}
    for tag in _COLOR_MOD_TAGS:
        child = color_el.find(f'{{{A}}}{tag}')
        if child is not None and child.get('val') is not None:
            mods[tag] = child.get('val')
    return mods


def parse_color(elem):
    """Extract color from a fill/stroke element.

    When the color has an explicit alpha channel, returns a dict like
    ``{'color': 'RRGGBB', 'alpha': 0.5}`` so transparency is preserved.
    """
    if elem is None:
        return None
    base = None
    color_el = None
    srgb = elem.find(f'{{{A}}}srgbClr')
    if srgb is not None:
        base = srgb.get('val')
        color_el = srgb
    scheme = elem.find(f'{{{A}}}schemeClr')
    if base is None and scheme is not None:
        base = f"scheme:{scheme.get('val')}"
        color_el = scheme
    scrgb = elem.find(f'{{{A}}}scrgbClr')
    if base is None and scrgb is not None:
        def c(v):
            return min(255, max(0, int(int(v or 0) * 255 / 100000)))
        base = f"{c(scrgb.get('r')):02X}{c(scrgb.get('g')):02X}{c(scrgb.get('b')):02X}"
        color_el = scrgb
    sysclr = elem.find(f'{{{A}}}sysClr')
    if base is None and sysclr is not None:
        last = sysclr.get('lastClr')
        base = last if last else f"sys:{sysclr.get('val')}"
        color_el = sysclr
    hsl = elem.find(f'{{{A}}}hslClr')
    if base is None and hsl is not None:
        base = f"hsl:{hsl.get('hue')},{hsl.get('sat')},{hsl.get('lum')}"
        color_el = hsl
    prst = elem.find(f'{{{A}}}prstClr')
    if base is None and prst is not None:
        base = f"prst:{prst.get('val')}"
        color_el = prst
    if base is None:
        return None
    alpha = _parse_alpha(color_el) if color_el is not None else None
    mods = _parse_color_mods(color_el) if color_el is not None else {}
    if (alpha is not None and alpha != 1.0) or mods:
        result = {'color': base}
        if alpha is not None and alpha != 1.0:
            result['alpha'] = alpha
        result.update(mods)
        return result
    return base


def parse_fill(spPr_or_tcPr):
    """Extract fill from shape properties or table cell properties.

    Returns a color string, 'none', 'image', or a dict for gradient/pattern fills.
    """
    if spPr_or_tcPr is None:
        return None
    solid = spPr_or_tcPr.find(f'{{{A}}}solidFill')
    if solid is not None:
        return parse_color(solid) or 'solid'
    grad = spPr_or_tcPr.find(f'{{{A}}}gradFill')
    if grad is not None:
        stops = []
        for gs in grad.iter(f'{{{A}}}gs'):
            pos = gs.get('pos')
            color = parse_color(gs)
            if pos is not None:
                stops.append({'pos': int(pos) / 100000, 'color': color})
        angle = 0
        lin = grad.find(f'{{{A}}}lin')
        if lin is not None:
            ooxml_angle = int(lin.get('ang', 0)) / 60000
            # OOXML angle is clockwise from 12 o'clock; python-pptx gradient_angle
            # is counter-clockwise. Convert to the value python-pptx would write.
            angle = (-ooxml_angle) % 360
        return {'type': 'gradient', 'angle': angle, 'stops': stops}
    patt = spPr_or_tcPr.find(f'{{{A}}}pattFill')
    if patt is not None:
        fg = parse_color(patt)
        bg = None
        bg_el = patt.find(f'{{{A}}}bgClr')
        if bg_el is not None:
            bg = parse_color(bg_el)
        return {
            'type': 'pattern',
            'name': patt.get('prst'),
            'fg': fg,
            'bg': bg,
        }
    noFill = spPr_or_tcPr.find(f'{{{A}}}noFill')
    if noFill is not None:
        return 'none'
    blipFill = spPr_or_tcPr.find(f'{{{A}}}blipFill')
    if blipFill is not None:
        return 'image'
    return None


def parse_line(ln_elem):
    """Parse an a:ln element into a dict."""
    if ln_elem is None:
        return None
    result = {}
    w = ln_elem.get('w')
    if w:
        result['w'] = round(int(w) / EMU_PER_POINT, 2)
    fill_elem = ln_elem.find(f'{{{A}}}solidFill')
    if fill_elem is not None:
        color = parse_color(fill_elem)
        if color:
            result['color'] = color
    noFill = ln_elem.find(f'{{{A}}}noFill')
    if noFill is not None:
        result['color'] = 'none'
    dash = ln_elem.find(f'{{{A}}}prstDash')
    if dash is not None:
        result['dash'] = dash.get('val')
    head = ln_elem.find(f'{{{A}}}headEnd')
    if head is not None:
        end = {}
        if head.get('type'):
            end['type'] = head.get('type')
        if head.get('w'):
            end['w'] = head.get('w')
        if head.get('len'):
            end['len'] = head.get('len')
        # A headEnd element with no explicit type but with size attributes denotes
        # the default arrow head in PowerPoint.
        if not end.get('type') and (end.get('w') or end.get('len')):
            end['type'] = 'arrow'
        result['head'] = end if end else 'arrow'
    tail = ln_elem.find(f'{{{A}}}tailEnd')
    if tail is not None:
        end = {}
        if tail.get('type'):
            end['type'] = tail.get('type')
        if tail.get('w'):
            end['w'] = tail.get('w')
        if tail.get('len'):
            end['len'] = tail.get('len')
        if not end.get('type') and (end.get('w') or end.get('len')):
            end['type'] = 'arrow'
        result['tail'] = end if end else 'arrow'
    cap = ln_elem.get('cap')
    if cap:
        result['cap'] = cap
    cmpd = ln_elem.get('cmpd')
    if cmpd:
        result['cmpd'] = cmpd
    return result if result else None


def parse_shape_line(spPr):
    """Extract line/border from shape properties."""
    if spPr is None:
        return None
    ln = spPr.find(f'{{{A}}}ln')
    if ln is None:
        return None
    return parse_line(ln)


def parse_effects(elem):
    """Extract text/shape effects (shadow, glow, reflection, etc.) as dicts."""
    if elem is None:
        return []
    effects = []
    for shdw in elem.iter(f'{{{A}}}outerShdw'):
        eff = {'type': 'outerShdw', 'color': parse_color(shdw) or '000000'}
        for attr in ('blurRad', 'dist', 'dir', 'algn', 'rotWithShape'):
            val = shdw.get(attr)
            if val is not None:
                eff[attr] = int(val) if attr != 'algn' else val
        effects.append(eff)
    for shdw in elem.iter(f'{{{A}}}innerShdw'):
        eff = {'type': 'innerShdw', 'color': parse_color(shdw) or '000000'}
        for attr in ('blurRad', 'dist', 'dir', 'algn'):
            val = shdw.get(attr)
            if val is not None:
                eff[attr] = int(val) if attr != 'algn' else val
        effects.append(eff)
    for glow in elem.iter(f'{{{A}}}glow'):
        eff = {'type': 'glow', 'color': parse_color(glow) or 'FFFFFF'}
        rad = glow.get('rad')
        eff['rad'] = int(rad) if rad is not None else 40000
        effects.append(eff)
    for refl in elem.iter(f'{{{A}}}reflection'):
        eff = {'type': 'reflection'}
        for attr in ('blurRad', 'dist', 'algn', 'rotWithShape'):
            val = refl.get(attr)
            if val is not None:
                eff[attr] = int(val) if attr != 'algn' else val
        effects.append(eff)
    for se in elem.iter(f'{{{A}}}softEdge'):
        rad = se.get('rad')
        effects.append({'type': 'softEdge', 'rad': int(rad) if rad is not None else 50000})
    return effects
def _hex_color(value):
    """Return a '#RRGGBB' string for SVG, or None if not resolvable."""
    if value is None:
        return None
    if isinstance(value, dict):
        value = value.get('color') or value.get('fill')
    if isinstance(value, str):
        if re.match(r'^[0-9A-Fa-f]{6}$', value):
            return f"#{value}"
        if value.startswith('#') and len(value) == 7:
            return value
    return None


def _svg_path_commands(path_elem, path_w, path_h):
    """Convert an <a:path> element's commands to an SVG path 'd' string.

    Supports moveTo, lnTo, cubicBezTo, quadBezTo, and close.  arcTo is
    approximated with a line to the end point because a full elliptical-arc
    conversion is overkill for the freeforms this skill normally encounters.
    Guide-name references in point coordinates are not evaluated; if a point
    is non-numeric the conversion fails and the caller should fall back to a
    placeholder.
    """
    def pt_val(cmd, idx=0):
        pts = cmd.findall(f'{{{A}}}pt')
        if not (0 <= idx < len(pts)):
            return None
        x = pts[idx].get('x')
        y = pts[idx].get('y')
        if x is None or y is None or not x.lstrip('-').isdigit() or not y.lstrip('-').isdigit():
            return None
        return int(x), int(y)

    def fmt(pt):
        if pt is None:
            return None
        # Keep path-space coordinates; add_image() scales to the slide position.
        return f"{pt[0]:.2f} {pt[1]:.2f}"

    parts = []
    for cmd in path_elem:
        tag = cmd.tag.split('}')[-1]
        if tag == 'moveTo':
            v = fmt(pt_val(cmd))
            if v is None:
                return None
            parts.append(f"M {v}")
        elif tag == 'lnTo':
            v = fmt(pt_val(cmd))
            if v is None:
                return None
            parts.append(f"L {v}")
        elif tag == 'quadBezTo':
            p1 = fmt(pt_val(cmd, 0))
            p2 = fmt(pt_val(cmd, 1))
            if p1 is None or p2 is None:
                return None
            parts.append(f"Q {p1} {p2}")
        elif tag == 'cubicBezTo':
            p1 = fmt(pt_val(cmd, 0))
            p2 = fmt(pt_val(cmd, 1))
            p3 = fmt(pt_val(cmd, 2))
            if p1 is None or p2 is None or p3 is None:
                return None
            parts.append(f"C {p1} {p2} {p3}")
        elif tag == 'arcTo':
            # Approximate with a line to the arc's end point.
            end = fmt(pt_val(cmd, 0))
            if end is None:
                return None
            parts.append(f"L {end}")
        elif tag == 'close':
            parts.append("Z")
    return " ".join(parts)


# Preset dash -> dash/gap pattern in multiples of the stroke width, matching
# PowerPoint's rendering. Consumed by _svg_dasharray to build stroke-dasharray.
_DASH_PATTERNS = {
    'dot': [1, 3],
    'sysDot': [1, 1],
    'dash': [4, 3],
    'sysDash': [3, 1],
    'lgDash': [8, 3],
    'dashDot': [4, 3, 1, 3],
    'sysDashDot': [3, 1, 1, 1],
    'lgDashDot': [8, 3, 1, 3],
    'lgDashDotDot': [8, 3, 1, 3, 1, 3],
    'sysDashDotDot': [3, 1, 1, 1, 1, 1],
}


def _svg_dasharray(dash_val, stroke_units):
    """Return an SVG stroke-dasharray for a preset dash, or None.

    ``stroke_units`` is the stroke width already expressed in the path
    coordinate space; OOXML dash lengths are relative to line width, so the
    pattern is scaled by it.
    """
    pattern = _DASH_PATTERNS.get(dash_val)
    if not pattern or not stroke_units:
        return None
    return " ".join(f"{seg * stroke_units:.2f}" for seg in pattern)


def _custgeom_to_svg(spPr):
    """Convert a <a:custGeom> shape to a standalone SVG string.

    Returns the SVG text, or None if the geometry uses unsupported features
    (guide-name point references, etc.).
    """
    custGeom = spPr.find(f'{{{A}}}custGeom')
    if custGeom is None:
        return None
    pathLst = custGeom.find(f'{{{A}}}pathLst')
    if pathLst is None:
        return None

    # Fill / stroke styling.
    fill_color = _hex_color(parse_fill(spPr))
    fill = f'fill="{fill_color}"' if fill_color else 'fill="none"'
    if spPr.find(f'{{{A}}}noFill') is not None:
        fill = 'fill="none"'

    line = parse_shape_line(spPr)
    stroke_color = _hex_color(line)
    stroke = 'stroke="none"'
    stroke_width = ''
    stroke_dash = ''
    if stroke_color and line:
        stroke = f'stroke="{stroke_color}"'
        # parse_line() stores width in POINTS (w_emu / EMU_PER_POINT); convert
        # back to EMU before scaling into the EMU-based path space, or the
        # stroke comes out ~EMU_PER_POINT times too thin and vanishes on raster.
        w_pt = line.get('w') if isinstance(line, dict) else 0
        sw_units = 0
        if w_pt:
            xfrm = spPr.find(f'{{{A}}}xfrm')
            if xfrm is not None:
                ext = xfrm.find(f'{{{A}}}ext')
                if ext is not None:
                    cx = int(ext.get('cx', 1))
                    # Use the first path's width as the geometry coordinate space.
                    first_path = pathLst.find(f'{{{A}}}path')
                    pw = int(first_path.get('w', 1)) if first_path is not None else 1
                    if cx and pw:
                        sw_units = w_pt * EMU_PER_POINT * pw / cx
                        stroke_width = f' stroke-width="{sw_units:.2f}"'
        # Preserve the dashed/dotted border (otherwise it renders solid).
        dasharray = _svg_dasharray(
            line.get('dash') if isinstance(line, dict) else None, sw_units
        )
        if dasharray:
            stroke_dash = f' stroke-dasharray="{dasharray}"'

    # Build path elements.
    paths = []
    max_w, max_h = 1, 1
    for path in pathLst.findall(f'{{{A}}}path'):
        pw = int(path.get('w', '1'))
        ph = int(path.get('h', '1'))
        d = _svg_path_commands(path, pw, ph)
        if d is None:
            return None
        max_w = max(max_w, pw)
        max_h = max(max_h, ph)
        paths.append((pw, ph, d))

    path_elems = []
    for pw, ph, d in paths:
        path_elems.append(f'<path d="{d}" {fill} {stroke}{stroke_width}{stroke_dash}/>')

    svg = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {max_w} {max_h}" '
        f'width="{max_w}" height="{max_h}">\n'
        + "\n".join(path_elems)
        + '\n</svg>\n'
    )
    return svg


# Extension URI for PowerPoint's "sketched"/hand-drawn line style
# (ask:lineSketchStyleProps, office/drawing/2018/sketchyshapes).
_SKETCH_EXT_URI = "{C807C97D-BFC1-408E-A445-0C87EB9F89A2}"


def _prstGeom_to_dict(prstGeom):
    """Build a geometry dict from an <a:prstGeom> element."""
    geom = {'type': prstGeom.get('prst')}
    adjustments = []
    avLst = prstGeom.find(f'{{{A}}}avLst')
    if avLst is not None:
        for gd in avLst.findall(f'{{{A}}}gd'):
            adjustments.append((gd.get('name'), gd.get('fmla')))
    if adjustments:
        geom['adj'] = adjustments
    return geom


def _sketch_preset_geometry(spPr):
    """Recover a shape's real preset geometry when a "sketch" line style is on.

    PowerPoint's hand-drawn line style (ask:lineSketchStyleProps) replaces the
    live geometry with a wavy <a:custGeom> bezier path, but preserves the
    original <a:prstGeom> inside the sketch extension (which sits in the line's
    extLst). Returning that preset lets the shape round-trip as a crisp native
    shape (rect, arrow, ...) instead of a rasterized freeform whose dashed edges
    come out broken. The decorative hand-drawn wobble is intentionally dropped.
    """
    for ext in spPr.iter(f'{{{A}}}ext'):
        if ext.get('uri') != _SKETCH_EXT_URI:
            continue
        for prstGeom in ext.iter(f'{{{A}}}prstGeom'):
            return _prstGeom_to_dict(prstGeom)
    return None


def parse_geometry(spPr):
    """Parse the <a:prstGeom> or <a:custGeom> element of a shape."""
    if spPr is None:
        return None
    prstGeom = spPr.find(f'{{{A}}}prstGeom')
    if prstGeom is not None:
        return _prstGeom_to_dict(prstGeom)
    custGeom = spPr.find(f'{{{A}}}custGeom')
    if custGeom is not None:
        # A sketch/hand-drawn line turns the real preset into a wavy custGeom;
        # prefer the preset stashed in the sketch extension so the shape stays a
        # crisp native preset instead of a rasterized (and broken) freeform.
        sketch_geom = _sketch_preset_geometry(spPr)
        if sketch_geom is not None:
            return sketch_geom
        svg_data = _custgeom_to_svg(spPr)
        geom = {'type': 'custom'}
        if svg_data:
            geom['svg_data'] = svg_data
        return geom
    return None


def parse_style(shape_elem):
    """Parse the <p:style> references (lnRef/fillRef/effectRef/fontRef)."""
    if shape_elem is None:
        return None
    style = shape_elem.find(f'{{{P}}}style')
    if style is None:
        return None
    refs = {}
    for tag, key in [
        (f'{{{A}}}lnRef', 'lnRef'),
        (f'{{{A}}}fillRef', 'fillRef'),
        (f'{{{A}}}effectRef', 'effectRef'),
        (f'{{{A}}}fontRef', 'fontRef'),
    ]:
        el = style.find(tag)
        if el is not None:
            ref = {'idx': el.get('idx')}
            # parse_color captures any shade/tint/lum transforms into the color
            # dict, so no separate extraction is needed here.
            color = parse_color(el)
            if color is not None:
                ref['color'] = color
            refs[key] = ref
    return refs


# ---------------------------------------------------------------------------
# Text parsing
# ---------------------------------------------------------------------------

def _resolve_hyperlink(hlink, slide_rels):
    """Return a hyperlink URL from an <a:hlinkClick> element."""
    if hlink is None:
        return None
    action = hlink.get('action')
    if action:
        return action
    r_id = hlink.get(f'{{{R}}}id')
    if slide_rels and r_id in slide_rels:
        return slide_rels[r_id]['target']
    return 'linked'


def parse_run(r, slide_rels=None):
    """Extract text and font attributes from a single <a:r> or <a:fld> run."""
    t = r.find(f'{{{A}}}t')
    text = t.text if t is not None and t.text else ''
    rPr = r.find(f'{{{A}}}rPr')
    info = {'text': text}
    if rPr is not None:
        info['sz'] = rPr.get('sz')
        info['b'] = rPr.get('b')
        info['i'] = rPr.get('i')
        info['u'] = rPr.get('u')
        info['strike'] = rPr.get('strike')
        info['baseline'] = rPr.get('baseline')
        info['spc'] = rPr.get('spc')
        for font_tag, key in [
            (f'{{{A}}}latin', 'typeface'),
            (f'{{{A}}}ea', 'ea'),
            (f'{{{A}}}cs', 'cs'),
        ]:
            latin = rPr.find(font_tag)
            if latin is not None:
                info[key] = latin.get('typeface')
                info[f'{key}_pitchFamily'] = latin.get('pitchFamily')
                info[f'{key}_charset'] = latin.get('charset')
        solidFill = rPr.find(f'{{{A}}}solidFill')
        if solidFill is not None:
            color = parse_color(solidFill)
            if color is not None:
                info['color'] = color
        highlight = rPr.find(f'{{{A}}}highlight')
        if highlight is not None:
            info['highlight'] = parse_color(highlight)
        effects = parse_effects(rPr)
        if effects:
            info['effects'] = effects
        hlink = rPr.find(f'{{{A}}}hlinkClick')
        if hlink is not None:
            info['hyperlink'] = _resolve_hyperlink(hlink, slide_rels)
    return info


_RUN_FONT_ATTRS = ['sz', 'b', 'i', 'u', 'strike', 'baseline', 'spc',
                   'typeface', 'ea', 'cs', 'color', 'highlight', 'effects',
                   'hyperlink',
                   'typeface_pitchFamily', 'ea_pitchFamily', 'cs_pitchFamily',
                   'typeface_charset', 'ea_charset', 'cs_charset']


def parse_default_run_props(defRPr):
    """Extract font attributes from a <a:defRPr> (paragraph/list default)."""
    if defRPr is None:
        return {}
    info = {}
    for k in ['sz', 'b', 'i', 'u', 'strike', 'baseline', 'spc']:
        v = defRPr.get(k)
        if v is not None:
            info[k] = v
    for font_tag, key in [
        (f'{{{A}}}latin', 'typeface'),
        (f'{{{A}}}ea', 'ea'),
        (f'{{{A}}}cs', 'cs'),
    ]:
        latin = defRPr.find(font_tag)
        if latin is not None:
            info[key] = latin.get('typeface')
            info[f'{key}_pitchFamily'] = latin.get('pitchFamily')
            info[f'{key}_charset'] = latin.get('charset')
    solidFill = defRPr.find(f'{{{A}}}solidFill')
    if solidFill is not None:
        color = parse_color(solidFill)
        if color is not None:
            info['color'] = color
    highlight = defRPr.find(f'{{{A}}}highlight')
    if highlight is not None:
        info['highlight'] = parse_color(highlight)
    effects = parse_effects(defRPr)
    if effects:
        info['effects'] = effects
    return info


def _parse_para_level_props(lvlPr):
    """Extract paragraph-level properties from a <a:lvlNpPr> or <a:defPPr>."""
    if lvlPr is None:
        return {}
    props = {}
    algn = lvlPr.get('algn')
    if algn is not None:
        props['algn'] = algn
    for attr in ('marL', 'indent', 'defTabSz'):
        v = lvlPr.get(attr)
        if v is not None:
            props[attr] = int(v)

    for tag, key in [
        (f'{{{A}}}lnSpc', 'lnSpc'),
        (f'{{{A}}}spcBef', 'spaceBefore'),
        (f'{{{A}}}spcAft', 'spaceAfter'),
    ]:
        el = lvlPr.find(tag)
        if el is None:
            continue
        spcPct = el.find(f'{{{A}}}spcPct')
        if spcPct is not None:
            props[key] = spcPct.get('val')
        spcPts = el.find(f'{{{A}}}spcPts')
        if spcPts is not None:
            props[key] = f"{spcPts.get('val')}pts"

    if lvlPr.find(f'{{{A}}}buNone') is not None:
        props['bullet'] = 'none'
    else:
        buChar = lvlPr.find(f'{{{A}}}buChar')
        if buChar is not None:
            props['bullet'] = 'char'
            char = buChar.get('char')
            if char:
                props['bullet_char'] = char
        buAutoNum = lvlPr.find(f'{{{A}}}buAutoNum')
        if buAutoNum is not None:
            props['bullet'] = 'autoNum'
            typ = buAutoNum.get('type')
            if typ:
                props['bullet_type'] = typ
        buBlip = lvlPr.find(f'{{{A}}}buBlip')
        if buBlip is not None:
            # Picture bullets cannot be round-tripped; fall back to a dot.
            props['bullet'] = 'blip'

    # Bullet glyph formatting: color, size (absolute points or percent of the
    # run font), and font. Captured independently of the glyph kind so a bullet
    # renders at its authored size instead of the full text size.
    buClr = lvlPr.find(f'{{{A}}}buClr')
    if buClr is not None:
        col = parse_color(buClr)
        if col is not None:
            props['bullet_color'] = col
    buSzPts = lvlPr.find(f'{{{A}}}buSzPts')
    if buSzPts is not None and buSzPts.get('val'):
        props['bullet_size_pts'] = int(buSzPts.get('val')) / 100
    else:
        buSzPct = lvlPr.find(f'{{{A}}}buSzPct')
        if buSzPct is not None and buSzPct.get('val'):
            props['bullet_size_pct'] = int(buSzPct.get('val')) / 1000
    buFont = lvlPr.find(f'{{{A}}}buFont')
    if buFont is not None and buFont.get('typeface'):
        props['bullet_font'] = buFont.get('typeface')
    return props


def _paragraph_default_props(txBody, pPr, placeholder_run_defaults=None):
    """Merge placeholder, list-style, and paragraph default run properties."""
    defaults = {}
    level = 0
    if pPr is not None:
        lvl_attr = pPr.get('lvl')
        if lvl_attr is not None:
            level = int(lvl_attr)

    if placeholder_run_defaults:
        defaults.update(placeholder_run_defaults.get(level, {}))
        if None in placeholder_run_defaults:
            # Fill any props not set by the level-specific default from the
            # placeholder-level default.
            for k, v in placeholder_run_defaults[None].items():
                defaults.setdefault(k, v)

    if txBody is not None:
        lstStyle = txBody.find(f'{{{A}}}lstStyle')
        if lstStyle is not None:
            # OOXML lvl="0" maps to a:lvl1pPr; no lvl attribute also means level 0.
            lvl_tag = f'{{{A}}}lvl{level + 1}pPr'
            lvlPr = lstStyle.find(lvl_tag)
            if lvlPr is not None:
                defaults.update(parse_default_run_props(lvlPr.find(f'{{{A}}}defRPr')))
            if not defaults:
                defPPr = lstStyle.find(f'{{{A}}}defPPr')
                if defPPr is not None:
                    defaults.update(parse_default_run_props(defPPr.find(f'{{{A}}}defRPr')))
    if pPr is not None:
        defaults.update(parse_default_run_props(pPr.find(f'{{{A}}}defRPr')))
    return defaults


def _paragraph_inherited_props(txBody, pPr, placeholder_para_defaults=None):
    """Merge placeholder, list-style, and paragraph-level properties."""
    defaults = {}
    level = 0
    if pPr is not None:
        lvl_attr = pPr.get('lvl')
        if lvl_attr is not None:
            level = int(lvl_attr)

    if placeholder_para_defaults:
        # Only inherit bullet and line spacing from placeholder defaults.
        # Spacing, indentation, and alignment inherited from layouts/masters
        # often duplicate defaults and can trigger PowerPoint repair dialogs
        # when emitted on non-placeholder shapes.
        inherited_keys = ('bullet', 'bullet_char', 'bullet_type', 'lnSpc')
        for src in (placeholder_para_defaults.get(None), placeholder_para_defaults.get(level)):
            if src:
                for key in inherited_keys:
                    if key in src:
                        defaults[key] = src[key]

    if txBody is not None:
        lstStyle = txBody.find(f'{{{A}}}lstStyle')
        if lstStyle is not None:
            lvl_tag = f'{{{A}}}lvl{level + 1}pPr'
            lvlPr = lstStyle.find(lvl_tag)
            if lvlPr is not None:
                parsed = _parse_para_level_props(lvlPr)
                for k, v in parsed.items():
                    defaults.setdefault(k, v)
            if not defaults:
                defPPr = lstStyle.find(f'{{{A}}}defPPr')
                if defPPr is not None:
                    parsed = _parse_para_level_props(defPPr)
                    for k, v in parsed.items():
                        defaults.setdefault(k, v)

    if pPr is not None:
        # The paragraph's own <a:pPr> overrides inherited defaults.
        parsed = _parse_para_level_props(pPr)
        defaults.update(parsed)

    return defaults


def _merge_run_with_defaults(default, run):
    """Apply default font props to a run, letting explicit run props win."""
    merged = dict(default)
    merged['text'] = run.get('text', '')
    for k in _RUN_FONT_ATTRS:
        if k in ('effects', 'highlight'):
            if run.get(k):
                merged[k] = run[k]
        elif run.get(k) is not None:
            merged[k] = run[k]
    return merged


def _math_xml_string(node):
    """Serialize an m:oMath element to a string, including its children."""
    return ET.tostring(node, encoding='unicode')


def _extract_math_xml(node):
    """Return the Office Math element, preserving the a14:m wrapper when present."""
    return node


def parse_text_body(txBody, slide_rels=None, placeholder_run_defaults=None, placeholder_para_defaults=None):
    """Parse a txBody into a list of paragraph dicts."""
    paragraphs = []
    for p in txBody.iter(f'{{{A}}}p'):
        text_parts = []
        runs = []
        pPr = p.find(f'{{{A}}}pPr')
        defaults = _paragraph_default_props(txBody, pPr, placeholder_run_defaults)
        para_defaults = _paragraph_inherited_props(txBody, pPr, placeholder_para_defaults)
        for node in p:
            if node.tag == f'{{{A}}}r' or node.tag == f'{{{A}}}fld':
                run = _merge_run_with_defaults(defaults, parse_run(node, slide_rels))
                runs.append(run)
                text_parts.append(run['text'])
            elif node.tag == f'{{{M}}}oMath' or node.tag == f'{{{A14}}}m':
                omath = _extract_math_xml(node)
                runs.append({'math_xml': _math_xml_string(omath)})
            elif node.tag == f'{{{A}}}br':
                text_parts.append('\n')
                runs.append({'text': '\n'})
        para = {
            'text': ''.join(text_parts),
            'runs': runs,
            'algn': para_defaults.get('algn'),
            'lnSpc': para_defaults.get('lnSpc'),
            'spaceBefore': para_defaults.get('spaceBefore'),
            'spaceAfter': para_defaults.get('spaceAfter'),
            'indent': para_defaults.get('indent'),
            'marL': para_defaults.get('marL'),
            'bullet': para_defaults.get('bullet'),
        }
        for extra in ('bullet_char', 'bullet_type', 'bullet_color',
                      'bullet_size_pts', 'bullet_size_pct', 'bullet_font'):
            if extra in para_defaults:
                para[extra] = para_defaults[extra]
        paragraphs.append(para)
    return paragraphs


# ---------------------------------------------------------------------------
# Transform handling and group flattening
# ---------------------------------------------------------------------------

def parse_xfrm(xfrm):
    """Parse an a:xfrm into a transform dict."""
    if xfrm is None:
        return None
    off = xfrm.find(f'{{{A}}}off')
    ext = xfrm.find(f'{{{A}}}ext')
    chOff = xfrm.find(f'{{{A}}}chOff')
    chExt = xfrm.find(f'{{{A}}}chExt')

    def _get(attr):
        # Some writers emit unqualified attributes, others namespace-qualify them.
        val = xfrm.get(attr)
        if val is not None:
            return val
        return xfrm.get(f'{{{A}}}{attr}')

    return {
        'off': (int(off.get('x', 0)), int(off.get('y', 0))) if off is not None else (0, 0),
        'ext': (int(ext.get('cx', 0)), int(ext.get('cy', 0))) if ext is not None else (0, 0),
        'chOff': (int(chOff.get('x', 0)), int(chOff.get('y', 0))) if chOff is not None else (0, 0),
        'chExt': (int(chExt.get('cx', 0)), int(chExt.get('cy', 0))) if chExt is not None else (0, 0),
        'rot': xfrm.get('rot', '0'),
        'flipH': _get('flipH'),
        'flipV': _get('flipV'),
    }


def xfrm_to_box(xfrm):
    """Convert a transform dict to (x, y, w, h) in inches."""
    if xfrm is None:
        return 0.0, 0.0, 0.0, 0.0
    x, y = xfrm['off']
    w, h = xfrm['ext']
    return x / EMU_PER_INCH, y / EMU_PER_INCH, w / EMU_PER_INCH, h / EMU_PER_INCH


def combine_transforms(parent, child):
    """Combine a parent group transform with a child transform.

    The child's (unrotated) extent is scaled by the parent group's
    child-space-to-display scale ``ext / chExt``. When the child is rotated a
    quarter turn (90 deg or 270 deg), its local axes are swapped relative to
    the group, so a *non-uniform* group scale must be applied with the x/y
    factors swapped -- rotation and non-uniform scaling do not commute, and
    without the swap a rotated shape (e.g. a -90 deg banner) comes out the
    wrong shape and drifts off its intended position. Positions are derived
    from the box center so the mapping is correct regardless of rotation.
    """
    if parent is None:
        return child
    if child is None:
        return parent
    px, py = parent['off']
    pcx, pcy = parent['ext']
    chx, chy = parent['chOff']
    chcx, chcy = parent['chExt']
    scale_x = pcx / chcx if chcx else 1.0
    scale_y = pcy / chcy if chcy else 1.0
    cx, cy = child['off']
    ccx, ccy = child['ext']
    # Map the child box center through the parent scale/translate; deriving the
    # offset from the center keeps this correct for rotated shapes too.
    center_x = px + (cx + ccx / 2 - chx) * scale_x
    center_y = py + (cy + ccy / 2 - chy) * scale_y
    try:
        child_deg = (int(child.get('rot', '0') or 0) / 60000.0) % 180
    except (TypeError, ValueError):
        child_deg = 0.0
    if abs(child_deg - 90) < 0.5:
        # Quarter turn: group x-scale acts on the child's height, y-scale on
        # its width.
        new_ext = (ccx * scale_y, ccy * scale_x)
    else:
        new_ext = (ccx * scale_x, ccy * scale_y)
    new_off = (center_x - new_ext[0] / 2, center_y - new_ext[1] / 2)
    return {
        'off': new_off,
        'ext': new_ext,
        'chOff': child['chOff'],
        'chExt': child['chExt'],
        'rot': str(int(parent.get('rot', '0') or 0) + int(child.get('rot', '0') or 0)),
        'flipH': child.get('flipH') or parent.get('flipH'),
        'flipV': child.get('flipV') or parent.get('flipV'),
    }


def _group_relative_transform(grp_xfrm, display_ext=None):
    """Return a transform that expresses child offsets relative to the group.

    ``display_ext`` is the group's *rendered* extent on the slide, i.e. its raw
    ``ext`` after any scaling inherited from ancestor groups has been applied. It
    must be used (instead of the group's own ``ext``) when computing the
    child-space-to-display scale ``ext / chExt``; otherwise a group nested inside
    another *scaled* group loses that ancestor scale and its descendants are
    rendered too small. Falls back to the raw ``ext`` for a top-level group,
    where the two are identical.
    """
    if grp_xfrm is None:
        return None
    return {
        'off': (0, 0),
        'ext': display_ext if display_ext is not None else grp_xfrm['ext'],
        'chOff': grp_xfrm['chOff'],
        'chExt': grp_xfrm['chExt'],
        'rot': '0',
    }


# ---------------------------------------------------------------------------
# Element parsers
# ---------------------------------------------------------------------------

def _base_attrs(elem, transform, group_path, typename, name, elem_id=''):
    """Common attributes for all elements.

    ``elem`` is the shape's ``a:xfrm`` (not the shape itself), so the caller
    passes ``elem_id`` extracted from the shape's ``cNvPr`` separately.
    """
    abs_xfrm = combine_transforms(transform, parse_xfrm(elem))
    x, y, w, h = xfrm_to_box(abs_xfrm)
    return {
        'type': typename,
        'name': name,
        'id': elem_id,
        'x': round(x, 3),
        'y': round(y, 3),
        'w': round(w, 3),
        'h': round(h, 3),
        'rot': abs_xfrm.get('rot') if abs_xfrm else '0',
        'flipH': abs_xfrm.get('flipH') if abs_xfrm else None,
        'flipV': abs_xfrm.get('flipV') if abs_xfrm else None,
        'group_path': list(group_path),
    }


def _parse_bodyPr_element(bodyPr):
    """Extract text body margins, anchor, wrap, and autofit from an <a:bodyPr>."""
    if bodyPr is None:
        return {}
    result = {}
    margins = {}
    for k in ['lIns', 'rIns', 'tIns', 'bIns']:
        v = bodyPr.get(k)
        if v is not None:
            margins[k] = round(int(v) / EMU_PER_INCH, 3)
    if margins:
        result['margins'] = margins
    anchor = bodyPr.get('anchor')
    if anchor is not None:
        result['anchor'] = anchor
    wrap = bodyPr.get('wrap')
    if wrap is not None:
        result['wrap'] = wrap
    autofit = None
    if bodyPr.find(f'{{{A}}}spAutoFit') is not None:
        autofit = 'spAutoFit'
    elif bodyPr.find(f'{{{A}}}noAutofit') is not None:
        autofit = 'noAutofit'
    if autofit is not None:
        result['autofit'] = autofit
    return result


def _parse_bodyPr(txBody):
    """Extract text body margins, anchor, wrap, and autofit from <a:bodyPr>."""
    if txBody is None:
        shape_bodyPr = {}
    else:
        shape_bodyPr = _parse_bodyPr_element(txBody.find(f'{{{A}}}bodyPr'))
    return (
        shape_bodyPr.get('margins'),
        shape_bodyPr.get('anchor') or 't',
        shape_bodyPr.get('wrap'),
        shape_bodyPr.get('autofit'),
    )


def _add_text(elem_dict, txBody, slide_rels=None,
              placeholder_run_defaults=None, placeholder_para_defaults=None):
    if txBody is None:
        return
    paragraphs = parse_text_body(
        txBody, slide_rels,
        placeholder_run_defaults=placeholder_run_defaults,
        placeholder_para_defaults=placeholder_para_defaults,
    )
    if paragraphs and all(
        p['text'] == '' and not any(r.get('text') or r.get('math_xml') for r in p['runs'])
        for p in paragraphs
    ):
        paragraphs = []
    elem_dict['paragraphs'] = paragraphs
    elem_dict['text'] = '\n'.join(p['text'] for p in paragraphs)
    # Capture body properties whenever there is real content, including
    # math-only boxes whose text is carried in run math_xml (so elem_dict
    # ['text'] is ''). Keying off `text` alone dropped their margins/wrap/
    # autofit, so an Office-Math equation lost its zero insets and spAutoFit
    # and wrapped to a second line in a box sized for one.
    if paragraphs:
        margins, anchor, wrap, autofit = _parse_bodyPr(txBody)
        if margins:
            elem_dict['margins'] = margins
        if anchor:
            elem_dict['anchor'] = anchor
        if wrap is not None:
            elem_dict['wrap'] = wrap
        if autofit is not None:
            elem_dict['autofit'] = autofit


def parse_sp(sp, transform, group_path, slide_rels=None, layout_ph_map=None, master_ph_map=None,
             layout_text_map=None, master_text_map=None, master_tx_styles=None,
             layout_bodyPr_map=None, master_bodyPr_map=None):
    """Parse a shape element."""
    name = get_name(sp, 'Shape')
    ph = sp.find(f'{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph')
    xfrm = sp.find(f'{{{P}}}spPr/{{{A}}}xfrm')
    if xfrm is None and ph is not None:
        xfrm = _find_inherited_xfrm(ph, layout_ph_map or {}, master_ph_map or {})
    if xfrm is None:
        return None
    elem = _base_attrs(xfrm, transform, group_path, 'shape', name, get_id(sp))
    spPr = sp.find(f'{{{P}}}spPr')
    if spPr is None:
        return None

    if elem['w'] == 0 and elem['h'] == 0:
        return None

    placeholder_run_defaults = None
    placeholder_para_defaults = None
    placeholder_bodyPr = None
    if ph is not None:
        elem['placeholder'] = ph.get('type')
        inherited = _find_inherited_text_defaults(
            ph, layout_text_map or {}, master_text_map or {}
        )
        master_style = _master_style_for_placeholder(ph, master_tx_styles or {})

        run_defaults = {}
        para_defaults = {}
        # Master txStyles are the base; placeholder defaults override them.
        for lvl, props in master_style.items():
            if props:
                run_defaults[lvl] = dict(props.get('run', {}))
                para_defaults[lvl] = dict(props.get('para', {}))
        if inherited:
            for lvl, props in inherited.items():
                if props.get('run'):
                    run_defaults.setdefault(lvl, {})
                    run_defaults[lvl].update(props['run'])
                if props.get('para'):
                    para_defaults.setdefault(lvl, {})
                    para_defaults[lvl].update(props['para'])
        if run_defaults:
            placeholder_run_defaults = run_defaults
        if para_defaults:
            placeholder_para_defaults = para_defaults

        placeholder_bodyPr = _find_inherited_bodyPr(
            ph, layout_bodyPr_map or {}, master_bodyPr_map or {}
        )

    xfrm = spPr.find(f'{{{A}}}xfrm')
    if xfrm is not None:
        if xfrm.get('rot') and xfrm.get('rot') != '0':
            elem['rot'] = xfrm.get('rot')

    elem['fill'] = parse_fill(spPr)
    elem['line'] = parse_shape_line(spPr)
    effects = parse_effects(spPr)
    if effects:
        elem['effects'] = effects

    geom = parse_geometry(spPr)
    if geom is not None:
        elem['geom'] = geom

    style = parse_style(sp)
    if style is not None:
        elem['style'] = style

    txBody = sp.find(f'{{{P}}}txBody')
    _add_text(
        elem, txBody, slide_rels,
        placeholder_run_defaults=placeholder_run_defaults,
        placeholder_para_defaults=placeholder_para_defaults,
    )

    hlinkClick = sp.find(f'.//{{{A}}}hlinkClick')
    if hlinkClick is not None:
        elem['hyperlink'] = _resolve_hyperlink(hlinkClick, slide_rels)

    return elem


def load_slide_rels(slide_path):
    """Build a dict of rId -> {'target': ..., 'type': ...} for a slide."""
    rels_path = slide_path.replace('/slides/', '/slides/_rels/').replace('.xml', '.xml.rels')
    if not os.path.exists(rels_path):
        return {}
    try:
        tree = ET.parse(rels_path)
    except ET.ParseError:
        return {}
    root = tree.getroot()
    rels = {}
    slides_dir = os.path.dirname(slide_path)
    for rel in root.findall(f'{{{PKG_R}}}Relationship'):
        r_id = rel.get('Id')
        target = rel.get('Target', '')
        rtype = rel.get('Type', '')
        if not r_id:
            continue
        abs_target = target
        if not target.startswith('http') and not target.startswith('mailto:'):
            abs_target = os.path.normpath(os.path.join(slides_dir, target))
        rels[r_id] = {'target': abs_target, 'type': rtype}
    return rels


def load_slide_image_rels(slide_path):
    """Build a dict of rId -> absolute media path for a slide."""
    rels = load_slide_rels(slide_path)
    image_rels = {}
    for r_id, info in rels.items():
        if 'image' in info['type']:
            image_rels[r_id] = info['target']
    return image_rels


def load_slide_layout_path(slide_path):
    """Return the absolute path to the slide layout XML for a slide."""
    rels_path = str(slide_path).replace('/slides/', '/slides/_rels/').replace('.xml', '.xml.rels')
    if not os.path.exists(rels_path):
        return None
    try:
        tree = ET.parse(rels_path)
    except ET.ParseError:
        return None
    root = tree.getroot()
    for rel in root.findall(f'{{{PKG_R}}}Relationship'):
        if 'slideLayout' in rel.get('Type', ''):
            target = rel.get('Target', '')
            slides_dir = os.path.dirname(str(slide_path))
            return os.path.normpath(os.path.join(slides_dir, target))
    return None


def load_layout_master_path(layout_path):
    """Return the absolute path to the slide master XML for a layout."""
    if not layout_path:
        return None
    rels_path = str(layout_path).replace('/slideLayouts/', '/slideLayouts/_rels/').replace('.xml', '.xml.rels')
    if not os.path.exists(rels_path):
        return None
    try:
        tree = ET.parse(rels_path)
    except ET.ParseError:
        return None
    root = tree.getroot()
    for rel in root.findall(f'{{{PKG_R}}}Relationship'):
        if 'slideMaster' in rel.get('Type', ''):
            target = rel.get('Target', '')
            layouts_dir = os.path.dirname(str(layout_path))
            return os.path.normpath(os.path.join(layouts_dir, target))
    return None


def _placeholder_key(ph):
    """Return (type, idx) tuple for a placeholder element."""
    return (ph.get('type'), ph.get('idx'))


def _parse_placeholder_xfrm_map(xml_path):
    """Build a dict of (type, idx) -> xfrm element for placeholders in a layout/master."""
    if not xml_path or not os.path.exists(xml_path):
        return {}
    try:
        tree = ET.parse(xml_path)
    except ET.ParseError:
        return {}
    root = tree.getroot()
    result = {}
    for sp in root.iter(f'{{{P}}}sp'):
        ph = sp.find(f'{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph')
        if ph is None:
            continue
        xfrm = sp.find(f'{{{P}}}spPr/{{{A}}}xfrm')
        if xfrm is None:
            continue
        result[_placeholder_key(ph)] = xfrm
    return result


def _parse_placeholder_text_defaults(xml_path):
    """Build a dict of (type, idx) -> {level: {'run': ..., 'para': ...}}.

    Levels are 0-based integers (OOXML lvl="0" -> level 0). A level of None
    holds any <a:defPPr> default.
    """
    if not xml_path or not os.path.exists(xml_path):
        return {}
    try:
        tree = ET.parse(xml_path)
    except ET.ParseError:
        return {}
    root = tree.getroot()
    result = {}
    for sp in root.iter(f'{{{P}}}sp'):
        ph = sp.find(f'{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph')
        if ph is None:
            continue
        txBody = sp.find(f'{{{P}}}txBody')
        if txBody is None:
            continue
        lstStyle = txBody.find(f'{{{A}}}lstStyle')
        if lstStyle is None:
            continue
        levels = {}
        defPPr = lstStyle.find(f'{{{A}}}defPPr')
        if defPPr is not None:
            levels[None] = {
                'run': parse_default_run_props(defPPr.find(f'{{{A}}}defRPr')),
                'para': _parse_para_level_props(defPPr),
            }
        for lvl in range(1, 10):
            lvlPr = lstStyle.find(f'{{{A}}}lvl{lvl}pPr')
            if lvlPr is None:
                continue
            levels[lvl - 1] = {
                'run': parse_default_run_props(lvlPr.find(f'{{{A}}}defRPr')),
                'para': _parse_para_level_props(lvlPr),
            }
        if levels:
            result[_placeholder_key(ph)] = levels
    return result


def _parse_master_tx_styles(master_path):
    """Return title/body default run/para props from the slide master txStyles.

    Result shape: {'title': {level: {'run': ..., 'para': ...}, ...}, 'body': {...}}.
    """
    result = {'title': {}, 'body': {}}
    if not master_path or not os.path.exists(master_path):
        return result
    try:
        root = ET.parse(master_path).getroot()
    except ET.ParseError:
        return result
    txStyles = root.find(f'{{{P}}}txStyles')
    if txStyles is None:
        return result
    for style_tag, key in [(f'{{{P}}}titleStyle', 'title'), (f'{{{P}}}bodyStyle', 'body')]:
        style_el = txStyles.find(style_tag)
        if style_el is None:
            continue
        for lvl in range(1, 10):
            lvlPr = style_el.find(f'{{{A}}}lvl{lvl}pPr')
            if lvlPr is None:
                continue
            result[key][lvl - 1] = {
                'run': parse_default_run_props(lvlPr.find(f'{{{A}}}defRPr')),
                'para': _parse_para_level_props(lvlPr),
            }
    return result


def _find_inherited_text_defaults(ph, layout_map, master_map):
    """Find placeholder default text props from layout or master."""
    key = _placeholder_key(ph)
    defaults = layout_map.get(key) or master_map.get(key)
    if defaults is not None:
        return defaults
    type_ = ph.get('type')
    if type_:
        defaults = layout_map.get((type_, None)) or master_map.get((type_, None))
        if defaults is not None:
            return defaults
    idx = ph.get('idx')
    if idx:
        defaults = layout_map.get((None, idx)) or master_map.get((None, idx))
    return defaults


def _master_style_for_placeholder(ph, tx_styles):
    """Return the master txStyles defaults relevant to a placeholder type."""
    type_ = ph.get('type')
    if type_ in ('title', 'ctrTitle', 'subTitle'):
        return tx_styles.get('title', {})
    # body, obj, text, and any other placeholder default to body style.
    return tx_styles.get('body', {})


def _parse_placeholder_bodyPr_map(xml_path):
    """Build a dict of (type, idx) -> bodyPr dict for placeholders."""
    if not xml_path or not os.path.exists(xml_path):
        return {}
    try:
        tree = ET.parse(xml_path)
    except ET.ParseError:
        return {}
    root = tree.getroot()
    result = {}
    for sp in root.iter(f'{{{P}}}sp'):
        ph = sp.find(f'{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph')
        if ph is None:
            continue
        txBody = sp.find(f'{{{P}}}txBody')
        if txBody is None:
            continue
        bodyPr = txBody.find(f'{{{A}}}bodyPr')
        if bodyPr is None:
            continue
        parsed = _parse_bodyPr_element(bodyPr)
        if parsed:
            result[_placeholder_key(ph)] = parsed
    return result


def _find_inherited_bodyPr(ph, layout_map, master_map):
    """Find placeholder bodyPr defaults from layout or master."""
    key = _placeholder_key(ph)
    defaults = layout_map.get(key) or master_map.get(key)
    if defaults is not None:
        return defaults
    type_ = ph.get('type')
    if type_:
        defaults = layout_map.get((type_, None)) or master_map.get((type_, None))
        if defaults is not None:
            return defaults
    idx = ph.get('idx')
    if idx:
        defaults = layout_map.get((None, idx)) or master_map.get((None, idx))
    return defaults


def _find_inherited_xfrm(ph, layout_map, master_map):
    """Find an xfrm for a placeholder inherited from layout or master."""
    key = _placeholder_key(ph)
    # Exact match
    xfrm = layout_map.get(key) or master_map.get(key)
    if xfrm is not None:
        return xfrm
    # Match by type only
    type_ = ph.get('type')
    if type_:
        xfrm = layout_map.get((type_, None)) or master_map.get((type_, None))
        if xfrm is not None:
            return xfrm
    # Match by idx only
    idx = ph.get('idx')
    if idx:
        xfrm = layout_map.get((None, idx)) or master_map.get((None, idx))
    return xfrm


def md5_file(path):
    if not path or not os.path.exists(path):
        return None
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def parse_pic(pic, transform, group_path, image_rels=None, slide_rels=None):
    """Parse an image or video element."""
    cNvPr = pic.find(f'{{{P}}}nvPicPr/{{{P}}}cNvPr')
    name = cNvPr.get('name', 'Picture') if cNvPr is not None else 'Picture'
    descr = cNvPr.get('descr', '') if cNvPr is not None else ''

    xfrm = pic.find(f'{{{P}}}spPr/{{{A}}}xfrm')
    if xfrm is None:
        return None
    elem = _base_attrs(xfrm, transform, group_path, 'image', name, get_id(pic))
    if elem['w'] == 0 and elem['h'] == 0:
        return None
    elem['descr'] = descr
    elem['text'] = descr.split('/')[-1] if descr else ''

    # Detect video/movie shapes.
    nvPr = pic.find(f'{{{P}}}nvPicPr/{{{P}}}nvPr')
    if nvPr is not None:
        videoFile = nvPr.find(f'{{{A}}}videoFile')
        if videoFile is not None:
            r_id = videoFile.get(f'{{{R}}}link')
            media_path = None
            if slide_rels and r_id:
                media_path = slide_rels[r_id]['target']
            elem['type'] = 'movie'
            elem['mediaFile'] = os.path.basename(media_path) if media_path else None

    spPr = pic.find(f'{{{P}}}spPr')
    if spPr is not None:
        elem['line'] = parse_shape_line(spPr)

    blipFill = pic.find(f'{{{P}}}blipFill')
    if blipFill is not None:
        blip = blipFill.find(f'{{{A}}}blip')
        img_info = {}
        if blip is not None:
            r_id = blip.get(f'{{{R}}}embed')
            # SVG-only pictures have no raster embed on <a:blip>; the SVG is
            # referenced via <asvg:svgBlip r:embed="..."> inside <a:extLst>.
            if not r_id:
                svg_blip = blip.find(
                    f'.//{{http://schemas.microsoft.com/office/drawing/2016/SVG/main}}svgBlip'
                )
                if svg_blip is not None:
                    r_id = svg_blip.get(f'{{{R}}}embed')
            if image_rels and r_id:
                media_path = image_rels.get(r_id)
                if media_path:
                    elem['imgHash'] = md5_file(media_path)
                    elem['imgFile'] = os.path.basename(media_path)
                    if elem.get('type') == 'movie':
                        elem['posterFile'] = os.path.basename(media_path)
            lum = blip.find(f'{{{A}}}lum')
            if lum is not None:
                img_info['lum'] = {
                    'bright': lum.get('bright'),
                    'contrast': lum.get('contrast'),
                }
        srcRect = blipFill.find(f'{{{A}}}srcRect')
        if srcRect is not None:
            img_info['crop'] = {k: srcRect.get(k) for k in ['l', 't', 'r', 'b']}
        if blipFill.find(f'{{{A}}}stretch') is not None:
            img_info['mode'] = 'stretch'
        if blipFill.find(f'{{{A}}}tile') is not None:
            img_info['mode'] = 'tile'
        if img_info:
            elem['imgInfo'] = img_info

    return elem


def parse_cxnSp(cxn, transform, group_path, slide_rels=None):
    """Parse a connector/line element."""
    name = get_name(cxn, 'Connector')
    xfrm = cxn.find(f'{{{P}}}spPr/{{{A}}}xfrm')
    if xfrm is None:
        return None
    elem = _base_attrs(xfrm, transform, group_path, 'connector', name, get_id(cxn))
    if elem['w'] == 0 and elem['h'] == 0:
        return None
    spPr = cxn.find(f'{{{P}}}spPr')
    if spPr is not None:
        elem['fill'] = parse_fill(spPr)
        elem['line'] = parse_shape_line(spPr)
        effects = parse_effects(spPr)
        if effects:
            elem['effects'] = effects
        geom = parse_geometry(spPr)
        if geom is not None:
            elem['geom'] = geom
        style = parse_style(cxn)
        if style is not None:
            elem['style'] = style
    # Shape-to-shape connections (stCxn/endCxn). These drive how a bent/elbow
    # connector routes its bend, so they must be preserved for fidelity.
    cxnPr = cxn.find(f'{{{P}}}nvCxnSpPr/{{{P}}}cNvCxnSpPr')
    if cxnPr is not None:
        connections = {}
        for side, tag in (('begin', 'stCxn'), ('end', 'endCxn')):
            node = cxnPr.find(f'{{{A}}}{tag}')
            if node is not None and node.get('id') is not None:
                connections[side] = {'id': node.get('id'), 'idx': node.get('idx', '0')}
        if connections:
            elem['connections'] = connections
    txBody = cxn.find(f'{{{P}}}txBody')
    _add_text(elem, txBody, slide_rels)
    return elem


def parse_cell(tc, slide_rels=None):
    """Parse a table cell."""
    txBody = tc.find(f'{{{A}}}txBody')
    paragraphs = parse_text_body(txBody, slide_rels) if txBody is not None else []
    text = '\n'.join(p['text'] for p in paragraphs)
    tcPr = tc.find(f'{{{A}}}tcPr')
    fill = None
    anchor = None
    margins = {}
    borders = {}
    grid_span = None
    row_span = None
    if tcPr is not None:
        fill = parse_fill(tcPr)
        anchor = tcPr.get('anchor')
        for k in ['marL', 'marR', 'marT', 'marB']:
            margins[k] = tcPr.get(k)
        for side, tag in [
            ('top', 'lnT'), ('right', 'lnR'),
            ('bottom', 'lnB'), ('left', 'lnL'),
        ]:
            ln = tcPr.find(f'{{{A}}}{tag}')
            if ln is not None:
                borders[side] = parse_line(ln)
        grid_span = tcPr.get('gridSpan')
        row_span = tcPr.get('rowSpan')
    return {
        'text': text,
        'paragraphs': paragraphs,
        'fill': fill,
        'anchor': anchor,
        'margins': margins,
        'borders': borders,
        'gridSpan': grid_span,
        'rowSpan': row_span,
    }


def parse_table(gf, transform, group_path, slide_rels=None):
    """Parse a table inside a graphicFrame."""
    abs_xfrm = combine_transforms(transform, parse_xfrm(gf.find(f'{{{P}}}xfrm')))
    x, y, w, h = xfrm_to_box(abs_xfrm)

    tbl = gf.find(f'.//{{{A}}}tbl')
    if tbl is None:
        return {'type': 'table', 'name': 'Table', 'x': x, 'y': y, 'w': w, 'h': h,
                'rows': 0, 'cols': 0, 'cells': [], 'colWidths': [], 'rowHeights': []}

    grid = tbl.find(f'{{{A}}}tblGrid')
    col_widths = []
    if grid is not None:
        for gc in grid.findall(f'{{{A}}}gridCol'):
            col_widths.append(round(int(gc.get('w', 0)) / EMU_PER_INCH, 3))

    rows = tbl.findall(f'{{{A}}}tr')
    cells = []
    row_heights = []
    for r in rows:
        row_heights.append(round(int(r.get('h', 0)) / EMU_PER_INCH, 3))
        row_cells = [parse_cell(tc, slide_rels) for tc in r.findall(f'{{{A}}}tc')]
        cells.append(row_cells)

    return {
        'type': 'table',
        'name': 'Table',
        'x': x, 'y': y, 'w': w, 'h': h,
        'rows': len(rows),
        'cols': len(col_widths),
        'colWidths': col_widths,
        'rowHeights': row_heights,
        'cells': cells,
    }


def _chart_type_name(chart_root):
    """Map a parsed chart XML root to an XL_CHART_TYPE enum name string."""
    plot_area = chart_root.find(f'{{{C}}}plotArea')
    if plot_area is None:
        return None
    for child in plot_area:
        tag = child.tag.split('}')[-1]
        if tag.endswith('Chart'):
            subtype = tag[:-5]
            # Determine direction/grouping.
            bar_dir = child.find(f'{{{C}}}barDir')
            grouping = child.find(f'{{{C}}}grouping')
            dir_part = ''
            group_part = 'CLUSTERED'
            if bar_dir is not None:
                dir_text = bar_dir.get('val', '')
                if dir_text == 'bar':
                    dir_part = 'BAR'
                elif dir_text == 'col':
                    dir_part = 'COLUMN'
            else:
                dir_part = subtype.upper()
            if grouping is not None:
                gval = grouping.get('val', '')
                if gval == 'stacked':
                    group_part = 'STACKED'
                elif gval == 'percentStacked':
                    group_part = 'PERCENT_STACKED'
            return f"{dir_part}_{group_part}"
    return None


def _chart_text(el):
    """Extract text from a chart title/series name element."""
    if el is None:
        return ''
    return child_text(el, C, 'v')


def _chart_categories(ser):
    """Return category labels from a c:ser/c:cat element."""
    cat = ser.find(f'{{{C}}}cat')
    if cat is None:
        return []
    pts = cat.findall(f'.//{{{C}}}pt')
    if pts:
        return [_chart_text(pt) for pt in pts]
    # Fall back to c:lit.
    lit = cat.find(f'{{{C}}}literal')
    if lit is not None:
        return [_chart_text(pt) for pt in lit.findall(f'{{{C}}}pt')]
    return []


def _chart_values(ser):
    """Return numeric values from a c:ser/c:val element."""
    val = ser.find(f'{{{C}}}val')
    if val is None:
        return []
    pts = val.findall(f'.//{{{C}}}pt')
    values = []
    for pt in pts:
        v = pt.find(f'{{{C}}}v')
        if v is not None and v.text:
            try:
                values.append(float(v.text))
            except ValueError:
                values.append(v.text)
        else:
            values.append(0)
    return values


def parse_chart_xml(chart_path: str) -> dict:
    """Parse a chart XML file into data the generator can rebuild."""
    if not chart_path or not os.path.exists(chart_path):
        return {}
    try:
        root = ET.parse(chart_path).getroot()
    except ET.ParseError:
        return {}

    chart_type = _chart_type_name(root)
    title = ''
    title_el = root.find(f'.//{{{C}}}chartTitle')
    if title_el is None:
        title_el = root.find(f'.//{{{C}}}title')
    if title_el is not None:
        title = _chart_text(title_el)

    plot_area = root.find(f'{{{C}}}plotArea')
    categories = []
    series = []
    if plot_area is not None:
        for ser in plot_area.findall(f'.//{{{C}}}ser'):
            if not categories:
                categories = _chart_categories(ser)
            name_el = ser.find(f'{{{C}}}tx')
            name = _chart_text(name_el)
            values = _chart_values(ser)
            series.append({'name': name, 'values': tuple(values)})

    return {
        'chart_type': chart_type,
        'title': title,
        'categories': categories,
        'series': series,
    }


def parse_chart_or_diagram(gf, transform, group_path, uri, slide_rels=None, chart_path=None):
    """Parse a chart or diagram graphicFrame."""
    abs_xfrm = combine_transforms(transform, parse_xfrm(gf.find(f'{{{P}}}xfrm')))
    x, y, w, h = xfrm_to_box(abs_xfrm)
    if 'chart' in uri:
        result = {'type': 'chart', 'name': 'Chart', 'subtype': 'chart',
                  'x': x, 'y': y, 'w': w, 'h': h, 'group_path': list(group_path),
                  'title': '', 'categories': [], 'series': [], 'chart_type': None}
        if chart_path:
            result.update(parse_chart_xml(chart_path))
        return result
    elif 'diagram' in uri:
        return {'type': 'diagram', 'name': 'Diagram', 'subtype': 'diagram',
                'x': x, 'y': y, 'w': w, 'h': h, 'group_path': list(group_path)}
    else:
        return {'type': 'graphic', 'name': 'Graphic', 'subtype': uri,
                'x': x, 'y': y, 'w': w, 'h': h, 'group_path': list(group_path)}


def parse_graphicFrame(gf, transform, group_path, slide_rels=None):
    """Parse a graphicFrame (table, chart, diagram, etc.)."""
    graphic = gf.find(f'.//{{{A}}}graphic')
    uri = ''
    r_id = None
    if graphic is not None:
        gd = graphic.find(f'{{{A}}}graphicData')
        if gd is not None:
            uri = gd.get('uri', '')
            r_id = gd.get(f'{{{R}}}id')

    chart_path = None
    if 'chart' in uri and slide_rels and r_id in slide_rels:
        chart_path = slide_rels[r_id]['target']

    if 'table' in uri:
        return parse_table(gf, transform, group_path, slide_rels)
    elif 'chart' in uri or 'diagram' in uri:
        return parse_chart_or_diagram(gf, transform, group_path, uri, slide_rels=slide_rels, chart_path=chart_path)
    else:
        abs_xfrm = combine_transforms(transform, parse_xfrm(gf.find(f'{{{P}}}xfrm')))
        x, y, w, h = xfrm_to_box(abs_xfrm)
        return {'type': 'graphic', 'name': 'Graphic', 'subtype': uri,
                'x': x, 'y': y, 'w': w, 'h': h, 'group_path': list(group_path)}


# ---------------------------------------------------------------------------
# Recursive extraction
# ---------------------------------------------------------------------------

def _extract_container(container, transform, group_path, slide_rels=None, image_rels=None, z_counter=None,
                       layout_ph_map=None, master_ph_map=None,
                       layout_text_map=None, master_text_map=None, master_tx_styles=None,
                       layout_bodyPr_map=None, master_bodyPr_map=None):
    """Recursively extract elements from a spTree or grpSp container."""
    if z_counter is None:
        z_counter = [0]
    elements = []
    for child in container:
        tag = child.tag
        if tag == f'{{{P}}}sp':
            e = parse_sp(child, transform, group_path, slide_rels, layout_ph_map, master_ph_map,
                         layout_text_map, master_text_map, master_tx_styles,
                         layout_bodyPr_map, master_bodyPr_map)
            if e:
                e['z'] = z_counter[0]
                z_counter[0] += 1
                elements.append(e)
        elif tag == f'{{{P}}}pic':
            e = parse_pic(child, transform, group_path, image_rels, slide_rels)
            if e:
                e['z'] = z_counter[0]
                z_counter[0] += 1
                elements.append(e)
        elif tag == f'{{{P}}}graphicFrame':
            e = parse_graphicFrame(child, transform, group_path, slide_rels)
            if e:
                e['z'] = z_counter[0]
                z_counter[0] += 1
                elements.append(e)
        elif tag == f'{{{P}}}cxnSp':
            e = parse_cxnSp(child, transform, group_path, slide_rels)
            if e:
                e['z'] = z_counter[0]
                z_counter[0] += 1
                elements.append(e)
        elif tag == f'{{{P}}}grpSp':
            cNvPr = child.find(f'{{{P}}}nvGrpSpPr/{{{P}}}cNvPr')
            grp_name = cNvPr.get('name', 'Group') if cNvPr is not None else 'Group'
            grp_xfrm = parse_xfrm(child.find(f'{{{P}}}grpSpPr/{{{A}}}xfrm'))
            abs_xfrm = combine_transforms(transform, grp_xfrm)
            gx, gy, gw, gh = xfrm_to_box(abs_xfrm)
            # Use the group's *display* extent (raw ext scaled by any ancestor
            # group scaling) so descendants keep the scale inherited from
            # enclosing scaled groups.
            display_ext = abs_xfrm['ext'] if abs_xfrm else None
            rel_transform = _group_relative_transform(grp_xfrm, display_ext)
            children = _extract_container(child, rel_transform, [], slide_rels, image_rels, z_counter,
                                          layout_ph_map, master_ph_map,
                                          layout_text_map, master_text_map, master_tx_styles,
                                          layout_bodyPr_map, master_bodyPr_map)
            group_elem = {
                'type': 'group',
                'name': grp_name,
                'id': cNvPr.get('id', '') if cNvPr is not None else '',
                'x': gx,
                'y': gy,
                'w': gw,
                'h': gh,
                'group_path': list(group_path),
                'children': children,
                'z': z_counter[0],
            }
            z_counter[0] += 1
            elements.append(group_elem)
        elif tag == f'{{{MC}}}AlternateContent':
            # Use the preferred Office 2010+ choice, falling back to the legacy representation.
            choice = child.find(f'{{{MC}}}Choice') or child.find(f'{{{MC}}}Fallback')
            if choice is not None:
                elements.extend(_extract_container(
                    choice, transform, group_path, slide_rels, image_rels, z_counter,
                    layout_ph_map, master_ph_map,
                    layout_text_map, master_text_map, master_tx_styles,
                    layout_bodyPr_map, master_bodyPr_map
                ))
    return elements


def extract_elements(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    spTree = root.find(f'.//{{{P}}}spTree')
    container = spTree if spTree is not None else root
    slide_rels = load_slide_rels(xml_path)
    image_rels = load_slide_image_rels(xml_path)

    layout_path = load_slide_layout_path(xml_path)
    layout_ph_map = _parse_placeholder_xfrm_map(layout_path) if layout_path else {}
    layout_text_map = _parse_placeholder_text_defaults(layout_path) if layout_path else {}
    layout_bodyPr_map = _parse_placeholder_bodyPr_map(layout_path) if layout_path else {}
    master_path = load_layout_master_path(layout_path) if layout_path else None
    master_ph_map = _parse_placeholder_xfrm_map(master_path) if master_path else {}
    master_text_map = _parse_placeholder_text_defaults(master_path) if master_path else {}
    master_bodyPr_map = _parse_placeholder_bodyPr_map(master_path) if master_path else {}
    master_tx_styles = _parse_master_tx_styles(master_path) if master_path else {'title': {}, 'body': {}}

    return _extract_container(container, None, [], slide_rels, image_rels,
                              layout_ph_map=layout_ph_map, master_ph_map=master_ph_map,
                              layout_text_map=layout_text_map, master_text_map=master_text_map,
                              master_tx_styles=master_tx_styles,
                              layout_bodyPr_map=layout_bodyPr_map, master_bodyPr_map=master_bodyPr_map)


def read_slide_shapes(slide_xml_path: Path) -> list[dict]:
    """Return one dict per shape/image/connector/table/etc. on the slide."""
    return extract_elements(str(slide_xml_path))


def parse_slide_notes(slide_xml_path: Path) -> str | None:
    """Extract the speaker-notes text for a slide, if any.

    Only text inside the notes body placeholder is returned; slide-number,
    header, date, and slide-image placeholders are ignored.
    """
    rels_path = str(slide_xml_path).replace('/slides/', '/slides/_rels/').replace('.xml', '.xml.rels')
    if not os.path.exists(rels_path):
        return None
    try:
        tree = ET.parse(rels_path)
    except ET.ParseError:
        return None
    root = tree.getroot()
    notes_target = None
    for rel in root.findall(f'{{{PKG_R}}}Relationship'):
        if 'notesSlide' in rel.get('Type', ''):
            notes_target = rel.get('Target')
            break
    if not notes_target:
        return None
    notes_path = os.path.normpath(os.path.join(os.path.dirname(str(slide_xml_path)), notes_target))
    if not os.path.exists(notes_path):
        return None
    try:
        notes_tree = ET.parse(notes_path)
    except ET.ParseError:
        return None
    notes_root = notes_tree.getroot()

    # Find the notes body placeholder shape.
    body_shape = None
    cSld = notes_root.find(f'{{{P}}}cSld')
    if cSld is not None:
        spTree = cSld.find(f'{{{P}}}spTree')
        if spTree is not None:
            for sp in spTree.findall(f'{{{P}}}sp'):
                nvPr = sp.find(f'{{{P}}}nvSpPr/{{{P}}}nvPr')
                if nvPr is not None:
                    ph = nvPr.find(f'{{{P}}}ph')
                    if ph is not None and ph.get('type') == 'body':
                        body_shape = sp
                        break

    if body_shape is None:
        return None

    txBody = body_shape.find(f'{{{P}}}txBody')
    if txBody is None:
        return None

    # Collect text only from the body placeholder, preserving paragraph breaks.
    paragraphs = []
    for p in txBody.findall(f'{{{A}}}p'):
        para_texts = [t.text for t in p.iter(f'{{{A}}}t') if t.text]
        if para_texts:
            paragraphs.append(''.join(para_texts))

    notes_text = '\n'.join(paragraphs).strip()
    if not notes_text:
        return None

    # Some source decks append the slide number to the notes body text
    # (e.g. "...edge device.4" on slide 4). Strip that trailing artifact
    # when it matches the slide index derived from the filename.
    slide_match = re.search(r'slide(\d+)\.xml$', str(slide_xml_path))
    if slide_match:
        slide_idx = int(slide_match.group(1))
        trailing_digits = re.search(r'(\d+)$', notes_text)
        if trailing_digits and int(trailing_digits.group(1)) == slide_idx:
            notes_text = notes_text[: trailing_digits.start()].rstrip()

    return notes_text if notes_text else None


def parse_slide_hidden(slide_xml_path: Path) -> bool:
    """Return True if the slide is hidden (``<p:sld show="0">``).

    Slide visibility is stored as the optional ``show`` attribute on the
    slide part's root ``<p:sld>`` element. The attribute is absent on visible
    slides; only an explicit ``show="0"`` marks a slide hidden.
    """
    try:
        root = ET.parse(slide_xml_path).getroot()
    except ET.ParseError:
        return False
    return root.get('show') == '0'


def parse_background(slide_xml_path: Path) -> dict | None:
    """Parse the slide background fill (<p:bg>) into a comparable dict."""
    tree = ET.parse(slide_xml_path)
    root = tree.getroot()
    cSld = root.find(f'{{{P}}}cSld')
    if cSld is None:
        return None
    bg = cSld.find(f'{{{P}}}bg')
    if bg is None:
        return None
    bgPr = bg.find(f'{{{P}}}bgPr')
    if bgPr is not None:
        solid = bgPr.find(f'{{{A}}}solidFill')
        if solid is not None:
            color = parse_color(solid)
            return {'type': 'solid', 'color': color}
        grad = bgPr.find(f'{{{A}}}gradFill')
        if grad is not None:
            stops = []
            for gs in grad.iter(f'{{{A}}}gs'):
                pos = gs.get('pos')
                color = parse_color(gs)
                if pos is not None:
                    stops.append({'pos': int(pos) / 100000, 'color': color})
            angle = 0
            lin = grad.find(f'{{{A}}}lin')
            if lin is not None:
                ooxml_angle = int(lin.get('ang', 0)) / 60000
                angle = (-ooxml_angle) % 360
            return {'type': 'gradient', 'angle': angle, 'stops': stops}
        noFill = bgPr.find(f'{{{A}}}noFill')
        if noFill is not None:
            return {'type': 'none'}
    bgRef = bg.find(f'{{{P}}}bgRef')
    if bgRef is not None:
        idx = bgRef.get('idx')
        color = parse_color(bgRef)
        return {'type': 'ref', 'idx': idx, 'color': color}
    return None
