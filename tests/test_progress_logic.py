import unittest
from datetime import date

from task_management.common.progress_logic import ProgressSource, sync_progress_records
from task_management.common.operations import _ordered_task_ids, _progress_row_mapping


class ProgressLogicTests(unittest.TestCase):
    def test_weekly_progress_uses_current_task_row_order(self):
        tasks = [
            {"個人タスクID": "PT_00002"},
            {"個人タスクID": ""},
            {"個人タスクID": "PT_00001"},
        ]
        self.assertEqual(_ordered_task_ids(tasks, "個人タスクID"), ["PT_00002", "PT_00001"])

    def test_personal_progress_mapping_uses_task_master_fields(self):
        task_headers = {name: index for index, name in enumerate(("個人タスクID", "テーマ名", "タスク名", "状態"), start=1)}
        progress_headers = {name: index for index, name in enumerate(("テーマ", "個人タスクID", "タスク名", "状態"), start=1)}
        self.assertEqual(
            _progress_row_mapping("personal", task_headers, progress_headers),
            {"テーマ": "テーマ名", "個人タスクID": "個人タスクID", "タスク名": "タスク名", "状態": "状態"},
        )
    def test_update_and_add_use_task_week_key(self):
        existing = [{"ログID": "LOG_00002", "タスクID": "T1", "週": date(2026, 8, 24), "進捗本文": "旧"}]
        sources = [
            ProgressSource("T1", "2026-08-26", "更新", {"テーマID": "A"}),
            ProgressSource("T2", "2026-08-31", "追加", {"テーマID": "B"}),
        ]
        result = sync_progress_records(sources, existing, "タスクID")
        self.assertEqual((result.updated, result.added), (1, 1))
        self.assertEqual(result.records[1]["ログID"], "LOG_00003")

    def test_blank_is_not_deleted_without_permission(self):
        existing = [{"ログID": "LOG_00001", "個人タスクID": "P1", "週": date(2026, 8, 24), "進捗本文": "保持"}]
        result = sync_progress_records([ProgressSource("P1", "2026-08-24", "", {})], existing, "個人タスクID", False)
        self.assertEqual(len(result.records), 1)
        self.assertEqual(result.deleted, 0)

    def test_duplicate_key_keeps_latest_record(self):
        existing = [
            {"ログID": "LOG_00001", "個人タスクID": "P1", "週": date(2026, 8, 3), "進捗本文": "旧", "更新日時": "2026-08-10T10:00:00"},
            {"ログID": "LOG_00002", "個人タスクID": "P1", "週": date(2026, 8, 3), "進捗本文": "新", "更新日時": "2026-08-11T10:00:00"},
        ]
        result = sync_progress_records([], existing, "個人タスクID")
        self.assertEqual(result.deduplicated, 1)
        self.assertEqual(len(result.records), 1)
        self.assertEqual(result.records[0]["進捗本文"], "新")

    def test_blank_template_rows_are_not_kept(self):
        existing = [
            {"ログID": "LOG_00001", "個人タスクID": "P1", "週": date(2026, 8, 3), "進捗本文": "本文"},
            {"ログID": None, "個人タスクID": None, "週": None, "進捗本文": None},
        ]
        result = sync_progress_records([], existing, "個人タスクID")
        self.assertEqual(len(result.records), 1)


if __name__ == "__main__":
    unittest.main()
