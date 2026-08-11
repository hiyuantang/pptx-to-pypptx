import io
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from extract_lecture_assets import extract_lecture_assets
from extract_notes import generate_target_markdown
from helpers.lecture_assets import remove_edge_background, trim_transparent
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
        self.assertIn("## Slide 1: First Topic", markdown)
        self.assertIn("First paragraph.\nSecond paragraph.", markdown)
        self.assertIn("## Slide 2: Second Topic", markdown)
        self.assertIn("_No speaker notes._", markdown)
        self.assertNotIn("\n1\n", markdown)

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

    def test_read_slide_size(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target = Path(tmpdir) / "lecture.pptx"
            make_fixture_pptx(target)
            width, height = read_slide_size(target)
        self.assertAlmostEqual(width, 13.333333, places=5)
        self.assertAlmostEqual(height, 7.5, places=5)


if __name__ == "__main__":
    unittest.main()
