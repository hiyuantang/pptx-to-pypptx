"""Isolate selected native PowerPoint shapes for lecture-note assets."""

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from helpers.pptx_utils import EMU_PER_INCH, count_slides, read_slide_size


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

for prefix, namespace in (("p", P), ("a", A), ("r", R), ("mc", MC)):
    ET.register_namespace(prefix, namespace)


_GROUP_TAG = f"{{{P}}}grpSp"
_ALTERNATE_CONTENT_TAG = f"{{{MC}}}AlternateContent"
_STRUCTURAL_TAGS = {f"{{{P}}}nvGrpSpPr", f"{{{P}}}grpSpPr"}


def parse_shape_ids(value: str) -> list[int]:
    """Parse a comma-separated list of positive PowerPoint shape IDs."""
    result: list[int] = []
    seen: set[int] = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            shape_id = int(part)
        except ValueError as exc:
            raise ValueError(f"Invalid shape ID: {part!r}") from exc
        if shape_id <= 0:
            raise ValueError("Shape IDs must be positive integers")
        if shape_id not in seen:
            result.append(shape_id)
            seen.add(shape_id)
    if not result:
        raise ValueError("Provide at least one shape ID")
    return result


def _shape_identity(element: ET.Element) -> tuple[int, str] | None:
    identity = element.find(f".//{{{P}}}cNvPr")
    if identity is None:
        return None
    try:
        shape_id = int(identity.get("id", ""))
    except ValueError:
        return None
    return shape_id, identity.get("name", "")


def _shape_inventory(sp_tree: ET.Element) -> dict[int, str]:
    inventory: dict[int, str] = {}
    for identity in sp_tree.findall(f".//{{{P}}}cNvPr"):
        try:
            shape_id = int(identity.get("id", ""))
        except ValueError:
            continue
        # The non-visual group root commonly uses ID 1 and is not selectable.
        if shape_id > 1:
            inventory.setdefault(shape_id, identity.get("name", ""))
    return inventory


def _filter_shape(element: ET.Element, selected: set[int]) -> bool:
    """Keep a selected shape, or the pruned ancestors of selected children."""
    identity = _shape_identity(element)
    if identity is not None and identity[0] in selected:
        return True

    if element.tag == _GROUP_TAG:
        kept_child = False
        for child in list(element):
            if child.tag in _STRUCTURAL_TAGS:
                continue
            if _filter_shape(child, selected):
                kept_child = True
            else:
                element.remove(child)
        return kept_child

    if element.tag == _ALTERNATE_CONTENT_TAG:
        # Choice/Fallback branches represent the same object. Keep the wrapper
        # intact when any representation contains a selected shape ID.
        descendant_ids = set()
        for identity in element.findall(f".//{{{P}}}cNvPr"):
            try:
                descendant_ids.add(int(identity.get("id", "")))
            except ValueError:
                continue
        return bool(descendant_ids & selected)

    return False


def _background_shape(
    shape_id: int,
    width_emu: int,
    height_emu: int,
    color: tuple[int, int, int],
) -> ET.Element:
    color_hex = "".join(f"{channel:02X}" for channel in color)
    return ET.fromstring(
        f'''<p:sp xmlns:p="{P}" xmlns:a="{A}">
  <p:nvSpPr>
    <p:cNvPr id="{shape_id}" name="Lecture asset temporary background"/>
    <p:cNvSpPr/><p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
</p:sp>'''
    )


def isolate_slide_shapes(
    source: Path,
    slide: int,
    shape_ids: list[int],
    destination: Path,
    *,
    background: tuple[int, int, int] = (0, 0, 0),
) -> dict:
    """Write a scratch PPTX containing only the selected objects on one slide.

    The rest of the package is copied verbatim so pictures, charts, fonts, and
    other relationships remain resolvable. A temporary flat slide-sized shape
    is inserted behind the selection to cover inherited layout/master chrome;
    callers remove that known color after rendering.
    """
    source = Path(source)
    destination = Path(destination)
    if not source.exists():
        raise FileNotFoundError(f"Target PPTX not found: {source}")
    if source.resolve() == destination.resolve():
        raise ValueError("Shape isolation destination must differ from the source PPTX")
    total = count_slides(source)
    if not 1 <= slide <= total:
        raise ValueError(f"Slide {slide} is outside the deck's 1-{total} range")
    if not shape_ids:
        raise ValueError("Provide at least one shape ID")
    if any(not 0 <= channel <= 255 for channel in background):
        raise ValueError("Background color channels must be between 0 and 255")

    slide_member = f"ppt/slides/slide{slide}.xml"
    with zipfile.ZipFile(source, "r") as archive:
        try:
            root = ET.fromstring(archive.read(slide_member))
        except KeyError as exc:
            raise ValueError(f"PPTX is missing {slide_member}") from exc
    sp_tree = root.find(f".//{{{P}}}spTree")
    if sp_tree is None:
        raise ValueError(f"Slide {slide} has no shape tree")

    inventory = _shape_inventory(sp_tree)
    selected = set(shape_ids)
    missing = sorted(selected - set(inventory))
    if missing:
        available = ", ".join(
            f"{shape_id} ({name or 'unnamed'})" for shape_id, name in sorted(inventory.items())
        )
        raise ValueError(
            f"Shape ID(s) not found on slide {slide}: {', '.join(map(str, missing))}. "
            f"Available: {available or 'none'}"
        )

    for child in list(sp_tree):
        if child.tag in _STRUCTURAL_TAGS:
            continue
        if not _filter_shape(child, selected):
            sp_tree.remove(child)

    slide_width, slide_height = read_slide_size(source)
    background_id = max(inventory, default=1) + 1
    background_shape = _background_shape(
        background_id,
        round(slide_width * EMU_PER_INCH),
        round(slide_height * EMU_PER_INCH),
        background,
    )
    insert_at = 0
    for index, child in enumerate(list(sp_tree)):
        if child.tag in _STRUCTURAL_TAGS:
            insert_at = index + 1
        else:
            break
    sp_tree.insert(insert_at, background_shape)
    slide_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(
        destination, "w"
    ) as destination_zip:
        for member in source_zip.infolist():
            data = slide_xml if member.filename == slide_member else source_zip.read(member.filename)
            destination_zip.writestr(member, data)

    return {
        "slide": slide,
        "shape_ids": shape_ids,
        "shape_names": [inventory[shape_id] for shape_id in shape_ids],
        "background": "#" + "".join(f"{channel:02x}" for channel in background),
    }
