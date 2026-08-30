import unittest

from task_management.personal.task import assign_management_ids, assign_personal_task_ids
from task_management.team.task import assign_task_ids


class IdLogicTests(unittest.TestCase):
    def test_team_ids_use_level_prefix_and_keep_existing(self):
        rows = [
            {"タスクID": "L1_00003", "Level": 1},
            {"タスクID": "", "Level": 1},
            {"タスクID": None, "Level": 2},
        ]
        result = assign_task_ids(rows)
        self.assertEqual([(x.row, x.value) for x in result], [(4, "L1_00004"), (5, "L2_00001")])

    def test_personal_ids_are_monotonic(self):
        rows = [{"個人タスクID": "PT_00009", "タスク名": "既存"}, {"個人タスクID": "", "タスク名": "新規"}]
        self.assertEqual(assign_personal_task_ids(rows)[0].value, "PT_00010")

    def test_risk_prefix_is_selected(self):
        rows = [{"課題・リスクID": "", "種別": "リスク", "内容": "懸念"}]
        self.assertEqual(assign_management_ids(rows, "課題・リスクID", "ISSUE")[0].value, "RISK_00001")

    def test_duplicate_is_rejected_before_assignment(self):
        rows = [{"個人タスクID": "PT_00001", "タスク名": "A"}, {"個人タスクID": "PT_00001", "タスク名": "B"}]
        with self.assertRaises(ValueError):
            assign_personal_task_ids(rows)


if __name__ == "__main__":
    unittest.main()
