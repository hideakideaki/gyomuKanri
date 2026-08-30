import unittest
from datetime import date

from task_management.personal.views import schedule_rows, this_week_rows, today_rows
from task_management.team.theme import build_theme_summary


class ViewTests(unittest.TestCase):
    def test_personal_views_show_leaf_and_standalone_tasks(self):
        tasks = [
            {"個人タスクID": "P1", "Level": 1, "親タスクID": "", "タスク名": "親", "状態": "進行中"},
            {"個人タスクID": "P2", "Level": 2, "親タスクID": "P1", "タスク名": "子", "状態": "進行中"},
            {"個人タスクID": "P3", "Level": 1, "親タスクID": "", "タスク名": "単独", "状態": "進行中"},
        ]
        rows = today_rows(tasks, date(2026, 8, 26))
        self.assertEqual({row["個人タスクID"] for row in rows}, {"P2", "P3"})

    def test_personal_views_exclude_closed_tasks(self):
        tasks = [
            {"個人タスクID": "P1", "タスク名": "進行", "状態": "進行中", "優先度": "高", "期限": date(2026, 8, 27), "予定開始日": date(2026, 8, 26)},
            {"個人タスクID": "P2", "タスク名": "完了", "状態": "完了", "優先度": "高", "期限": date(2026, 8, 26)},
        ]
        self.assertEqual(len(this_week_rows(tasks, date(2026, 8, 24), date(2026, 8, 26))), 1)
        self.assertEqual(len(today_rows(tasks, date(2026, 8, 26))), 1)
        self.assertEqual(len(schedule_rows(tasks, date(2026, 8, 24))), 1)

    def test_theme_summary_counts_open_records(self):
        tasks = [
            {"タスクID": "L1", "テーマID": "T1", "Level": 1, "L1": "テーマ", "状態": "進行中"},
            {"タスクID": "L2", "テーマID": "T1", "Level": 2, "状態": "完了"},
        ]
        issues = [{"テーマID": "T1", "状態": "対応中"}]
        result = build_theme_summary(tasks, issues, {"T1": "保持"})[0]
        self.assertEqual(result["未完了タスク数"], 1)
        self.assertEqual(result["課題・リスク件数"], 1)
        self.assertEqual(result["テーマ総括"], "保持")


if __name__ == "__main__":
    unittest.main()
