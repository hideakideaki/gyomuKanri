import unittest
from datetime import date, datetime, timezone

from task_management.common.indexes import build_header_index, build_id_index, build_week_index, monday_iso


class IndexTests(unittest.TestCase):
    def test_header_index_survives_column_insertion(self):
        headers = ["タスクID", "追加列", "状態"]
        self.assertEqual(build_header_index(headers)["状態"], 3)

    def test_duplicate_header_is_rejected(self):
        with self.assertRaises(ValueError):
            build_header_index(["状態", "状態"])

    def test_id_index_survives_sort(self):
        index = build_id_index([(3, "T-2"), (4, "T-1")], "タスクID")
        self.assertEqual(index["T-1"], 4)

    def test_duplicate_id_is_rejected(self):
        with self.assertRaises(ValueError):
            build_id_index([(3, "T-1"), (7, "T-1")], "タスクID")

    def test_week_index_uses_monday(self):
        index = build_week_index([(9, date(2026, 8, 26)), (11, "2026/09/02週")])
        self.assertEqual(index, {"2026-08-24": 9, "2026-08-31": 11})

    def test_week_index_accepts_non_zero_padded_dates(self):
        self.assertEqual(monday_iso("2026/8/19週"), "2026-08-17")

    def test_excel_utc_value_is_converted_to_japan_date(self):
        excel_value = datetime(2026, 8, 9, 15, 0, tzinfo=timezone.utc)
        self.assertEqual(monday_iso(excel_value), "2026-08-10")


if __name__ == "__main__":
    unittest.main()
