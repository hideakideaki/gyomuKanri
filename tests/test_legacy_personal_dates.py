import unittest

from task_management.common.models import Record, TableData, WorkbookData
from task_management.common.service import _normalize_legacy_personal_dates
from task_management.common.specs import PERSONAL


class LegacyPersonalDateTests(unittest.TestCase):
    def test_legacy_execution_date_moves_to_planned_start(self):
        data = WorkbookData(
            PERSONAL.format,
            "2026-08-29T00:00:00",
            [TableData("01_個人タスク", "個人タスクID", [Record("P1", {"実施予定日": "2026-08-10", "実施予定週": "2026-08-03"})])],
            [],
        )
        _normalize_legacy_personal_dates(data, PERSONAL)
        fields = data.tables[0].records[0].fields
        self.assertEqual(fields["予定開始日"], "2026-08-10")
        self.assertNotIn("実施予定日", fields)
        self.assertNotIn("実施予定週", fields)

    def test_existing_planned_start_takes_priority(self):
        data = WorkbookData(
            PERSONAL.format,
            "2026-08-29T00:00:00",
            [TableData("01_個人タスク", "個人タスクID", [Record("P1", {"予定開始日": "2026-08-01", "実施予定日": "2026-08-10"})])],
            [],
        )
        _normalize_legacy_personal_dates(data, PERSONAL)
        self.assertEqual(data.tables[0].records[0].fields, {"予定開始日": "2026-08-01", "親タスクID": "", "Level": 1})


if __name__ == "__main__":
    unittest.main()
