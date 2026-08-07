import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

import ai8d


class Ai8DTests(unittest.TestCase):
    def test_numeric_prefix_images_share_one_page(self):
        images = [Path("1-1.png"), Path("1-2.png"), Path("2-1.png")]
        pages = ai8d.plan_image_pages(images, [], "")
        self.assertEqual([[item.name for item in page["images"]] for page in pages], [["1-1.png", "1-2.png"], ["2-1.png"]])

    def test_clear_targets_is_recursive_and_restricted_to_stage_folders(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            (root / "D0 相关信息" / "nested").mkdir(parents=True)
            (root / "D0 相关信息" / "nested" / "evidence.png").touch()
            (root / "D0backup").mkdir()
            (root / "D0backup" / "keep.txt").touch()
            self.assertEqual([item.name for item in ai8d.clear_targets(root)], ["evidence.png"])

    def test_current_template_can_be_located_without_slide_indexes(self):
        template = next(Path(__file__).resolve().parent.parent.glob("*.pptx"))
        pages = ai8d.locate_template_slides(Presentation(template))
        self.assertEqual(pages["d4_occurrence"], 7)
        self.assertEqual(pages["d8"], 12)


if __name__ == "__main__":
    unittest.main()
