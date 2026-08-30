import tempfile
import unittest
from pathlib import Path

from task_management.common.json_io import canonical_payload, dump_v3, load_v3
from task_management.common.models import Record, RichTextRun, RichTextValue, TableData, WeeklyProgress, WorkbookData


class JsonTests(unittest.TestCase):
    def sample(self):
        return WorkbookData(
            "HierarchicalTaskManager-3",
            "2026-08-26T12:00:00",
            [TableData("02_タスク", "タスクID", [Record("L3_00001", {"タスク名": "方式検討", "工数": 4.5, "完了": False})])],
            [WeeklyProgress("L3_00001", "2026-08-24", RichTextValue("AAA\nBBB\n日本語CCC", [RichTextRun(1, 4, "#000000"), RichTextRun(5, 3, "#FF0000", True), RichTextRun(8, 7, "#000000")]))],
        )

    def test_round_trip_and_utf8_readability(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "data.json"
            source = self.sample()
            dump_v3(source, path)
            text = path.read_text(encoding="utf-8")
            self.assertIn("方式検討", text)
            self.assertNotIn("Base64", text)
            self.assertEqual(canonical_payload(source), canonical_payload(load_v3(path)))

    def test_invalid_rich_text_range_is_rejected(self):
        data = self.sample()
        data.weekly_progress[0].value.runs = [RichTextRun(10, 99, "#000000")]
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError):
                dump_v3(data, Path(directory) / "bad.json")


if __name__ == "__main__":
    unittest.main()
