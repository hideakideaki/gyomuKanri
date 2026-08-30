import unittest

from task_management.common.import_plan import create_plan
from task_management.common.models import Record, RichTextValue, TableData, WeeklyProgress, WorkbookData
from task_management.common.service import _excel_import_value, _planned_additions
from task_management.common.specs import PERSONAL


class ImportPlanTests(unittest.TestCase):
    def test_id_header_and_week_indexes_are_direct(self):
        data = WorkbookData(
            "HierarchicalTaskManager-3",
            "2026-08-26T00:00:00",
            [TableData("02_タスク", "タスクID", [Record("T1", {"状態": "完了"})])],
            [WeeklyProgress("T1", "2026-08-24", RichTextValue("進捗"))],
        )
        indexes = {
            "02_タスク": {"headers": {"タスクID": 1, "状態": 8}, "rows": {"T1": 30}},
            "05_週次進捗": {"headers": {"タスクID": 2}, "rows": {"T1": 12}},
        }
        plan = create_plan(data, indexes, {"05_週次進捗": {"2026-08-24": 18}}, "05_週次進捗")
        self.assertTrue(plan.valid)
        self.assertEqual((plan.updates[0].row, plan.updates[0].column), (30, 8))
        self.assertEqual((plan.updates[1].row, plan.updates[1].column), (12, 18))

    def test_missing_id_and_header_do_not_guess(self):
        data = WorkbookData("PersonalTaskManager-Keyed-3", "x", [TableData("01_個人タスク", "個人タスクID", [Record("P9", {"状態": "進行中"})])], [])
        plan = create_plan(data, {"01_個人タスク": {"headers": {"状況": 3}, "rows": {}}}, {}, "05_週次振り返り")
        self.assertFalse(plan.valid)
        self.assertTrue(plan.missing_ids)

    def test_missing_items_are_planned_as_explicit_additions(self):
        data = WorkbookData(
            "PersonalTaskManager-Keyed-3",
            "x",
            [TableData("01_個人タスク", "個人タスクID", [Record("P9", {"個人タスクID": "P9", "旧版列": "値"})])],
            [WeeklyProgress("P9", "2025-01-06", RichTextValue("進捗"))],
        )
        indexes = {
            "01_個人タスク": {"headers": {"個人タスクID": 1}, "rows": {}},
            "05_週次振り返り": {"headers": {"個人タスクID": 2}, "rows": {}},
        }
        rows, columns, weeks = _planned_additions(data, indexes, {"05_週次振り返り": {}}, PERSONAL)
        self.assertEqual(rows, [("01_個人タスク", "P9"), ("05_週次振り返り", "P9")])
        self.assertEqual(columns, [("01_個人タスク", "旧版列")])
        self.assertEqual(weeks, ["2025-01-06"])

    def test_iso_utc_datetime_is_restored_as_japan_excel_datetime(self):
        value = _excel_import_value("2026-08-16T15:00:00+00:00", "yyyy/m/d")
        self.assertEqual(str(value), "2026-08-17 00:00:00")

    def test_iso_like_text_stays_text_in_general_cell(self):
        self.assertEqual(_excel_import_value("2026-08-17", "General"), "2026-08-17")


if __name__ == "__main__":
    unittest.main()
