import unittest
from datetime import date

from task_management.common.operations import EXCEL_BUSY_HRESULT, _gantt_color, _gantt_rows, _set_gantt_fill, _task_period


class _Interior:
    def __init__(self, failures=0, hresult=EXCEL_BUSY_HRESULT):
        self.failures = failures
        self.hresult = hresult
        self.values = []

    @property
    def Color(self):
        return self.values[-1] if self.values else None

    @Color.setter
    def Color(self, value):
        if self.failures:
            self.failures -= 1
            error = RuntimeError("Excel busy")
            error.hresult = self.hresult
            raise error
        self.values.append(value)


class _Target:
    def __init__(self, interior):
        self.Interior = interior


class GanttLogicTests(unittest.TestCase):

    def test_gantt_fill_retries_excel_busy_error(self):
        interior = _Interior(failures=2)
        _set_gantt_fill(_Target(interior), 123, attempts=3)
        self.assertEqual(interior.values, [123])

    def test_gantt_fill_does_not_hide_other_errors(self):
        interior = _Interior(failures=1, hresult=-1)
        with self.assertRaises(RuntimeError):
            _set_gantt_fill(_Target(interior), 123)
    def test_personal_period_fallback(self):
        row = {"予定開始日": date(2026, 8, 26), "期限": None}
        self.assertEqual(_task_period(row, "personal"), (date(2026, 8, 26), date(2026, 8, 26)))

    def test_team_weekly_limits_to_level_two(self):
        tasks = [{"タスクID": "A", "Level": 1}, {"タスクID": "B", "Level": 3}]
        self.assertEqual(len(_gantt_rows(tasks, "team", True)), 1)
        self.assertEqual(len(_gantt_rows(tasks, "team", False)), 2)

    def test_team_normal_color_depends_on_level(self):
        today = date(2026, 8, 30)
        end = date(2026, 9, 30)
        self.assertEqual(_gantt_color({"Level": 1, "状態": "進行中"}, "team", end, today), "#2F5597")
        self.assertEqual(_gantt_color({"Level": 2, "状態": "未着手"}, "team", end, today), "#5B9BD5")
        self.assertEqual(_gantt_color({"Level": 3, "状態": "未着手"}, "team", end, today), "#9DC3E6")
        self.assertEqual(_gantt_color({"Level": 4, "状態": "未着手"}, "team", end, today), "#DDEBF7")

    def test_personal_parent_and_child_use_different_colors(self):
        today, end = date(2026, 8, 30), date(2026, 9, 1)
        self.assertEqual(_gantt_color({"Level": 1, "状態": "進行中"}, "personal", end, today), "#2F5597")
        self.assertEqual(_gantt_color({"Level": 2, "状態": "進行中"}, "personal", end, today), "#5B9BD5")

    def test_alert_color_takes_priority_over_level(self):
        today = date(2026, 8, 30)
        self.assertEqual(_gantt_color({"Level": 1, "状態": "待ち"}, "team", date(2026, 9, 1), today), "#F6B26B")
        self.assertEqual(_gantt_color({"Level": 2, "状態": "進行中"}, "team", date(2026, 8, 29), today), "#E06666")
        self.assertEqual(_gantt_color({"Level": 3, "状態": "完了"}, "team", date(2026, 8, 29), today), "#A6A6A6")


if __name__ == "__main__":
    unittest.main()
