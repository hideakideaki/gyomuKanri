import unittest

from task_management.team.task import validate_hierarchy


class HierarchyTests(unittest.TestCase):
    def test_valid_hierarchy(self):
        rows = [
            {"タスクID": "L1_00001", "Level": 1, "親タスクID": "", "テーマID": "THEME_00001"},
            {"タスクID": "L2_00001", "Level": 2, "親タスクID": "L1_00001", "テーマID": "THEME_00001"},
        ]
        self.assertEqual(validate_hierarchy(rows), [])

    def test_parent_level_and_theme_are_checked(self):
        rows = [
            {"タスクID": "L1_00001", "Level": 1, "親タスクID": "", "テーマID": "A"},
            {"タスクID": "L3_00001", "Level": 3, "親タスクID": "L1_00001", "テーマID": "B"},
        ]
        errors = validate_hierarchy(rows)
        self.assertTrue(any("親Level" in error for error in errors))
        self.assertTrue(any("テーマID" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
