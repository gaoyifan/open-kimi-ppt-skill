#!/usr/bin/env python3
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pptx import Presentation


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))
SCRIPT = SCRIPTS / "export_pptx_local.py"
SPEC = importlib.util.spec_from_file_location("export_pptx_local", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


class ExportPptxLocalTests(unittest.TestCase):
    def test_exports_rendered_pages_as_flattened_slides(self):
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            manifest = root / "deck.pptd"
            manifest.write_text("title: Flattened deck\npages: []\n", encoding="utf-8")
            output = root / "deck.pptx"

            def fake_export_images(_manifest, image_output):
                pages = image_output / "pages"
                pages.mkdir(parents=True)
                images = []
                for index, color in enumerate(("#ff0000", "#0000ff"), start=1):
                    path = pages / f"{index}.png"
                    Image.new("RGB", (1600, 900), color).save(path)
                    images.append({"image": f"pages/{index}.png"})
                return {"images": images}

            with patch.object(MODULE, "export_images", side_effect=fake_export_images):
                summary = MODULE.export_pptx(manifest, output, "fade", False)

            presentation = Presentation(output)
            self.assertEqual(len(presentation.slides), 2)
            self.assertEqual(presentation.core_properties.title, "Flattened deck")
            self.assertTrue(summary["flattened"])
            self.assertEqual(summary["transitionPatchedSlides"], 2)
            self.assertEqual(summary["slides"], 2)


if __name__ == "__main__":
    unittest.main()
