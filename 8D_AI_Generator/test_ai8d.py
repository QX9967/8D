import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

import ai8d


class Ai8DTests(unittest.TestCase):
    def test_large_material_prompt_uses_text_evidence_without_raw_images(self):
        materials = {
            "D4 原因分析": {
                "notes": "对异常样件进行复核。",
                "images": [Path("1-1.png"), Path("1-2.png")],
            }
        }
        prompt = ai8d.build_prompt(
            materials,
            include_images=False,
            image_evidence={"D4 原因分析": [{"image": "1-1.png", "summary": "已复核装配状态。"}]},
        )
        self.assertTrue(any("图片分批排查结论" in item["text"] for item in prompt))
        self.assertFalse(any(item["type"] == "image_url" for item in prompt))

    def test_split_batches_keeps_request_size_bounded(self):
        batches = ai8d.split_batches(list(range(23)), 10)
        self.assertEqual([len(batch) for batch in batches], [10, 10, 3])
        self.assertEqual(sum(batches, []), list(range(23)))

    def test_numeric_prefix_images_share_one_page(self):
        images = [Path("1-1.png"), Path("1-2.png"), Path("2-1.png")]
        pages = ai8d.plan_image_pages(images, [{"images": ["1-1.png"], "summary": "第一张证据"}], "")
        self.assertEqual([[item.name for item in page["images"]] for page in pages], [["1-1.png", "1-2.png"], ["2-1.png"]])
        self.assertEqual(pages[0]["summary"], "第一张证据；相关现场证据见图。")
        self.assertEqual(pages[0]["image_summaries"], ["第一张证据", "相关现场证据见图。"])

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
