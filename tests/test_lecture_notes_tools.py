import io
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from extract_lecture_assets import extract_lecture_assets
from extract_notes import generate_target_markdown
from finalize_lecture_notes import finalize_markdown
from helpers.lecture_assets import (
    recover_alpha_from_mattes,
    remove_edge_background,
    trim_transparent,
)
from helpers.lecture_shapes import isolate_slide_shapes, parse_shape_ids
from helpers.pptx_utils import read_slide_size
from prepare_lecture_asset import prepare_asset
from validate_lecture_notes import validate_lecture_notes


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"


def _slide_xml(title: str) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="914400"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{title}</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="3" name="Picture 3" descr="Transparent Diagram.png"/>
        <p:cNvPicPr/><p:nvPr/>
      </p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>
    <p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="4" name="Diagram Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm>
        <a:off x="3657600" y="1828800"/><a:ext cx="1828800" cy="914400"/>
        <a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/>
      </a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="Diagram Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="6" name="Unselected Callout"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
      </p:sp>
    </p:grpSp>
  </p:spTree></p:cSld>
</p:sld>'''


def _slide_rels(include_notes: bool) -> str:
    notes = (
        f'<Relationship Id="rIdNotes" Type="{R}/notesSlide" '
        'Target="../notesSlides/notesSlide1.xml"/>'
        if include_notes else ""
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_R}">
  <Relationship Id="rIdImage" Type="{R}/image" Target="../media/image1.png"/>
  {notes}
</Relationships>'''


def _notes_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="{P}" xmlns:a="{A}">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:t>First paragraph.</a:t></a:r></a:p>
        <a:p><a:r><a:t>Second paragraph.</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Slide Number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr>
      <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>1</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>'''


def _transparent_png() -> bytes:
    image = Image.new("RGBA", (8, 6), (20, 40, 60, 255))
    image.putpixel((0, 0), (0, 0, 0, 0))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def make_fixture_pptx(path: Path) -> None:
    presentation = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="{P}"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>'''
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ppt/presentation.xml", presentation)
        archive.writestr("ppt/slides/slide1.xml", _slide_xml("First Topic"))
        archive.writestr("ppt/slides/slide2.xml", _slide_xml("Second Topic"))
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", _slide_rels(include_notes=True))
        archive.writestr("ppt/slides/_rels/slide2.xml.rels", _slide_rels(include_notes=False))
        archive.writestr("ppt/notesSlides/notesSlide1.xml", _notes_xml())
        archive.writestr("ppt/media/image1.png", _transparent_png())


class LectureNotesToolTests(unittest.TestCase):
    def test_extract_notes_directly_from_pptx(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target = Path(tmpdir) / "lecture.pptx"
            make_fixture_pptx(target)
            markdown = generate_target_markdown(target)

        self.assertIn("# Speaker Notes: lecture", markdown)
        self.assertIn("<!-- lecture-source-slide: 1 -->", markdown)
        self.assertIn("## Slide 1: First Topic", markdown)
        self.assertIn("First paragraph.\nSecond paragraph.", markdown)
        self.assertIn("## Slide 2: Second Topic", markdown)
        self.assertIn("_No speaker notes._", markdown)
        self.assertNotIn("\n1\n", markdown)

    def test_extract_notes_links_temporary_previews_by_actual_slide_number(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target = Path(tmpdir) / "lecture.pptx"
            make_fixture_pptx(target)
            markdown = generate_target_markdown(
                target,
                selected_slides=[2],
                preview_paths={2: "slide-images/slide_2.png"},
            )

        self.assertNotIn("## Slide 1:", markdown)
        self.assertIn("<!-- lecture-source-slide: 2 -->", markdown)
        self.assertIn("## Slide 2: Second Topic", markdown)
        self.assertIn("(slide-images/slide_2.png)", markdown)
        self.assertIn("<!-- lecture-source-preview:start -->", markdown)

    def test_finalize_removes_preview_and_provenance_markers(self):
        draft = """# Lecture title

<!-- source-slides: 4-6 -->
## Core concept

<!-- lecture-source-preview:start -->
![Temporary full-slide reference for slide 4 — not a final lecture-note asset](slide-images/slide_4.png)
<!-- lecture-source-preview:end -->

Concept prose.
"""

        finalized, stats = finalize_markdown(draft)

        self.assertEqual(finalized, "# Lecture title\n\n## Core concept\n\nConcept prose.\n")
        self.assertEqual(stats["preview_blocks_removed"], 1)
        self.assertEqual(stats["provenance_markers_removed"], 1)

    def test_finalize_rejects_unedited_slide_structure(self):
        draft = "# Lecture title\n\n## Slide 4: Core concept\n\nSpeaker notes.\n"

        with self.assertRaisesRegex(ValueError, "slide-number headings"):
            finalize_markdown(draft)

    def test_isolate_slide_shapes_keeps_only_requested_object_and_background(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            target = root / "lecture.pptx"
            isolated = root / "isolated.pptx"
            make_fixture_pptx(target)

            result = isolate_slide_shapes(target, 1, [3], isolated)

            with zipfile.ZipFile(isolated, "r") as archive:
                slide_xml = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            identities = {
                int(node.get("id")): node.get("name")
                for node in slide_xml.findall(f".//{{{P}}}cNvPr")
            }

        self.assertEqual(result["shape_ids"], [3])
        self.assertIn(3, identities)
        self.assertEqual(identities[3], "Picture 3")
        self.assertNotIn(2, identities)
        self.assertIn("Lecture asset temporary background", identities.values())

    def test_shape_id_parser_deduplicates_and_validates(self):
        self.assertEqual(parse_shape_ids("7, 3,7"), [7, 3])
        with self.assertRaisesRegex(ValueError, "positive"):
            parse_shape_ids("0")

    def test_isolate_selected_group_child_prunes_its_siblings(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            target = root / "lecture.pptx"
            isolated = root / "isolated-child.pptx"
            make_fixture_pptx(target)

            isolate_slide_shapes(target, 1, [5], isolated)

            with zipfile.ZipFile(isolated, "r") as archive:
                slide_xml = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            identities = {
                int(node.get("id")): node.get("name")
                for node in slide_xml.findall(f".//{{{P}}}cNvPr")
            }

        self.assertIn(4, identities)
        self.assertEqual(identities[5], "Diagram Box")
        self.assertNotIn(6, identities)
        self.assertNotIn(2, identities)
        self.assertNotIn(3, identities)

    def test_extract_assets_preserves_alpha_and_deduplicates_usage(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            target = root / "lecture.pptx"
            output = root / "candidates"
            make_fixture_pptx(target)
            manifest = extract_lecture_assets(target, output, [1, 2])

            self.assertEqual(manifest["summary"]["extracted_assets"], 1)
            self.assertEqual(manifest["summary"]["resolved_picture_usages"], 2)
            self.assertEqual(manifest["summary"]["total_picture_usages"], 2)
            asset = manifest["assets"][0]
            self.assertTrue(asset["has_transparency"])
            self.assertEqual([usage["slide"] for usage in asset["usages"]], [1, 2])
            with Image.open(output / asset["file"]) as image:
                self.assertEqual(image.getpixel((0, 0))[3], 0)

    def test_remove_background_only_when_connected_to_edge(self):
        image = Image.new("RGBA", (7, 7), (255, 255, 255, 255))
        for x in range(2, 5):
            image.putpixel((x, 2), (0, 0, 0, 255))
            image.putpixel((x, 4), (0, 0, 0, 255))
        for y in range(2, 5):
            image.putpixel((2, y), (0, 0, 0, 255))
            image.putpixel((4, y), (0, 0, 0, 255))

        result, stats = remove_edge_background(image, tolerance=0)

        self.assertEqual(result.getpixel((0, 0))[3], 0)
        self.assertEqual(result.getpixel((3, 3)), (255, 255, 255, 255))
        self.assertEqual(result.getpixel((2, 2)), (0, 0, 0, 255))
        self.assertGreater(stats["removed_pixels"], 0)

    def test_background_inference_refuses_ambiguous_corners(self):
        image = Image.new("RGBA", (5, 5), (255, 255, 255, 255))
        image.putpixel((0, 0), (255, 0, 0, 255))
        image.putpixel((4, 0), (0, 255, 0, 255))
        image.putpixel((0, 4), (0, 0, 255, 255))
        image.putpixel((4, 4), (255, 255, 0, 255))

        with self.assertRaisesRegex(ValueError, "Cannot infer a uniform edge background"):
            remove_edge_background(image, tolerance=0)

    def test_background_inference_refuses_transparent_corners(self):
        image = Image.new("RGBA", (5, 5), (0, 0, 0, 0))
        image.putpixel((2, 2), (0, 0, 0, 255))

        with self.assertRaisesRegex(ValueError, "fewer than three opaque corners"):
            remove_edge_background(image, tolerance=0)

    def test_trim_preserves_semitransparent_edge_alpha(self):
        image = Image.new("RGBA", (3, 3), (0, 0, 0, 0))
        image.putpixel((1, 1), (20, 40, 60, 128))

        result = trim_transparent(image, padding=2)

        self.assertEqual(result.size, (5, 5))
        self.assertEqual(result.getpixel((2, 2)), (20, 40, 60, 128))

    def test_two_matte_recovery_preserves_antialiased_color_and_alpha(self):
        source = Image.new("RGBA", (3, 1), (0, 0, 0, 0))
        source.putpixel((1, 0), (60, 180, 120, 128))
        source.putpixel((2, 0), (10, 20, 30, 255))
        dark = Image.new("RGBA", source.size, (0, 0, 0, 255))
        light = Image.new("RGBA", source.size, (255, 255, 255, 255))
        dark.alpha_composite(source)
        light.alpha_composite(source)

        recovered = recover_alpha_from_mattes(dark, light)

        self.assertEqual(recovered.getpixel((0, 0))[3], 0)
        semitransparent = recovered.getpixel((1, 0))
        self.assertTrue(all(abs(actual - expected) <= 2 for actual, expected in zip(
            semitransparent, (60, 180, 120, 128)
        )))
        self.assertEqual(recovered.getpixel((2, 0)), (10, 20, 30, 255))

    def test_prepare_and_validate_transparent_asset(self):
        image = Image.new("RGBA", (9, 9), (255, 255, 255, 255))
        image.putpixel((4, 4), (0, 0, 0, 255))
        prepared, _ = prepare_asset(
            image,
            transparent=True,
            background=(255, 255, 255),
            tolerance=0,
            trim=True,
            padding=2,
        )
        self.assertEqual(prepared.size, (5, 5))

        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            prepared.save(assets / "concept.png")
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Lecture\n\n![Concept diagram](lecture-notes-assets/concept.png)\n",
                encoding="utf-8",
            )
            validation = validate_lecture_notes(markdown, assets)

        self.assertEqual(validation["errors"], [])
        self.assertEqual(validation["summary"]["linked_assets"], 1)
        self.assertEqual(validation["summary"]["opaque_rasters"], 0)

    def test_validator_warns_for_opaque_asset_by_default(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            Image.new("RGB", (8, 6), (20, 40, 60)).save(assets / "screenshot.png")
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Lecture\n\n![Application screenshot](lecture-notes-assets/screenshot.png)\n",
                encoding="utf-8",
            )

            validation = validate_lecture_notes(markdown, assets)

        self.assertEqual(validation["errors"], [])
        self.assertIn(
            "Opaque raster asset (inspect whether transparency is appropriate): screenshot.png",
            validation["warnings"],
        )

    def test_strict_transparency_rejects_unapproved_opaque_asset(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            Image.new("RGB", (8, 6), (20, 40, 60)).save(assets / "diagram.png")
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Lecture\n\n![Native diagram](lecture-notes-assets/diagram.png)\n",
                encoding="utf-8",
            )

            validation = validate_lecture_notes(
                markdown,
                assets,
                strict_transparency=True,
            )

        self.assertTrue(any(
            "Opaque raster asset is not explicitly allowed in strict mode: diagram.png"
            in error
            for error in validation["errors"]
        ))

    def test_strict_transparency_accepts_exact_opaque_allowlist(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            Image.new("RGB", (8, 6), (20, 40, 60)).save(assets / "screenshot.png")
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Lecture\n\n![Application screenshot](lecture-notes-assets/screenshot.png)\n",
                encoding="utf-8",
            )

            validation = validate_lecture_notes(
                markdown,
                assets,
                strict_transparency=True,
                allowed_opaque=["screenshot.png"],
            )

        self.assertEqual(validation["errors"], [])
        self.assertEqual(validation["warnings"], [])
        self.assertEqual(validation["summary"]["allowed_opaque_rasters"], 1)

    def test_strict_transparency_rejects_stale_opaque_allowlist_entry(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            Image.new("RGBA", (8, 6), (20, 40, 60, 0)).save(assets / "diagram.png")
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Lecture\n\n![Native diagram](lecture-notes-assets/diagram.png)\n",
                encoding="utf-8",
            )

            validation = validate_lecture_notes(
                markdown,
                assets,
                strict_transparency=True,
                allowed_opaque=["diagram.png"],
            )

        self.assertIn(
            "Allowed opaque asset is not a linked opaque raster: diagram.png",
            validation["errors"],
        )

    def test_validator_rejects_nonstandard_math_delimiters(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Lecture\n\nInline \\(x\\).\n\n\\[y = 2\\]\n",
                encoding="utf-8",
            )

            validation = validate_lecture_notes(markdown, assets)

        self.assertIn(
            "Use $...$ and $$...$$ for Markdown math, not \\(...\\) or \\[...\\]",
            validation["errors"],
        )

    def test_validator_rejects_temporary_full_slide_reference(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            Image.new("RGBA", (4, 4), (0, 0, 0, 0)).save(assets / "slide_2.png")
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Lecture\n\n![Slide reference](lecture-notes-assets/slide_2.png)\n",
                encoding="utf-8",
            )

            validation = validate_lecture_notes(markdown, assets)

        self.assertIn(
            "Temporary full-slide reference leaked into final lecture notes: "
            "lecture-notes-assets/slide_2.png",
            validation["errors"],
        )

    def test_validator_rejects_source_slide_structure_and_provenance(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            assets = root / "lecture-notes-assets"
            assets.mkdir()
            markdown = root / "lecture-notes.md"
            markdown.write_text(
                "# Speaker Notes: Lecture\n\n<!-- source-slides: 2-3 -->\n"
                "## Slide 2: Concept\n\nDraft prose.\n",
                encoding="utf-8",
            )

            validation = validate_lecture_notes(markdown, assets)

        self.assertIn(
            "Source-slide provenance remains; run finalize_lecture_notes.py",
            validation["errors"],
        )
        self.assertIn(
            "Slide-by-slide source structure remains; organize the final notes by concept",
            validation["errors"],
        )

    def test_read_slide_size(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target = Path(tmpdir) / "lecture.pptx"
            make_fixture_pptx(target)
            width, height = read_slide_size(target)
        self.assertAlmostEqual(width, 13.333333, places=5)
        self.assertAlmostEqual(height, 7.5, places=5)


if __name__ == "__main__":
    unittest.main()
