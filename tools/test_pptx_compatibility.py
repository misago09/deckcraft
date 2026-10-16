"""Check the independent fixture and comparison harness, not the product."""

import io
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

from PIL import Image

import pptx_compatibility as gate


class CompatibilityHarnessTests(unittest.TestCase):
    def test_fixture_has_valid_xml_and_resolvable_internal_relationships(self):
        with tempfile.TemporaryDirectory() as d:
            path = gate.fixtures(Path(d))
            with zipfile.ZipFile(path) as z:
                for name in z.namelist():
                    if name.endswith(('.xml', '.rels')):
                        ET.fromstring(z.read(name))
                image = Image.open(io.BytesIO(z.read('ppt/media/pattern.png')))
                self.assertEqual(image.size, (4, 4))
            sig = gate.signature(path)
            self.assertEqual(len(sig['slides']), 6)
            self.assertEqual(sig['broken_relationships'], [])
            self.assertEqual(sig['invalid_xml'], [])
            self.assertEqual(sig['slides'][0]['Latin and Hangul']['text'][1], '굵은 한글 123')
            self.assertEqual(sig['slides'][2]['Rotated group']['geometry']['rot'], '900000')

    def test_pixel_comparator_detects_movement_not_just_size(self):
        with tempfile.TemporaryDirectory() as d:
            a, b = Path(d)/'a.png', Path(d)/'b.png'
            image = Image.new('RGB', (4, 4), 'white')
            image.save(a)
            image.save(b)
            self.assertTrue(gate.compare_pixels(a,b)['pass'])
            image.putpixel((2,2), (0,0,0))
            image.save(b)
            changed = gate.compare_pixels(a,b)
            self.assertFalse(changed['pass'])
            self.assertEqual(changed['changed_pixels_over_8'],1)

    def test_xml_check_catches_undeclared_prefix_outside_slide_parts(self):
        with tempfile.TemporaryDirectory() as d:
            path = gate.fixtures(Path(d))
            with zipfile.ZipFile(path, 'a') as z:
                z.writestr('ppt/comments/invalid.xml', '<p188:cmLst xmlns:p188="urn:test"><a:p/></p188:cmLst>')
            invalid = gate.signature(path)['invalid_xml']
            self.assertEqual(len(invalid),1)
            self.assertEqual(invalid[0]['part'],'ppt/comments/invalid.xml')
            self.assertIn('unbound prefix',invalid[0]['error'])


if __name__ == '__main__':
    unittest.main()
