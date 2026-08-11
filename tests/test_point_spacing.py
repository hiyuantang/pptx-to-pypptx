import sys
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from helpers.slide_codegen import _paragraph_props
from helpers.slide_xml import A, _parse_para_level_props


class PointSpacingTests(unittest.TestCase):
    def test_drawingml_point_spacing_is_normalized_once(self):
        paragraph = ET.fromstring(
            f'''<a:lvl1pPr xmlns:a="{A}">
              <a:lnSpc><a:spcPts val="3000"/></a:lnSpc>
              <a:spcBef><a:spcPts val="1000"/></a:spcBef>
              <a:spcAft><a:spcPts val="1425"/></a:spcAft>
            </a:lvl1pPr>'''
        )

        parsed = _parse_para_level_props(paragraph)
        generated = _paragraph_props(parsed)

        self.assertEqual(parsed["lnSpc"], "30pts")
        self.assertEqual(parsed["spaceBefore"], "10pts")
        self.assertEqual(parsed["spaceAfter"], "14.25pts")
        self.assertEqual(generated["line_spacing"], "30pts")
        self.assertEqual(generated["space_before"], 10.0)
        self.assertEqual(generated["space_after"], 14.25)


if __name__ == "__main__":
    unittest.main()
