import time
import unittest

from task_management.common.import_plan import create_plan
from task_management.common.models import Record, TableData, WorkbookData


class PerformanceTests(unittest.TestCase):
    def test_5000_records_are_indexed_linearly(self):
        count = 5000
        records = [Record(f"T{i}", {"状態": "進行中"}) for i in range(count)]
        data = WorkbookData("HierarchicalTaskManager-3", "x", [TableData("02_タスク", "タスクID", records)], [])
        indexes = {"02_タスク": {"headers": {"状態": 8}, "rows": {f"T{i}": i + 3 for i in range(count)}}}
        started = time.perf_counter()
        plan = create_plan(data, indexes, {}, "05_週次進捗")
        elapsed = time.perf_counter() - started
        self.assertEqual(len(plan.updates), count)
        self.assertLess(elapsed, 2.0)


if __name__ == "__main__":
    unittest.main()
