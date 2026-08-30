import unittest

from task_management.personal.task import normalize_personal_hierarchy, validate_personal_hierarchy


class PersonalHierarchyTests(unittest.TestCase):
    def test_valid_parent_child_and_standalone(self):
        rows = [
            {"個人タスクID": "P1", "Level": 1, "親タスクID": ""},
            {"個人タスクID": "P2", "Level": 2, "親タスクID": "P1"},
            {"個人タスクID": "P3", "Level": 1, "親タスクID": ""},
        ]
        self.assertEqual(validate_personal_hierarchy(rows), [])

    def test_invalid_parent_is_reported(self):
        rows = [{"個人タスクID": "P2", "Level": 2, "親タスクID": "UNKNOWN"}]
        self.assertTrue(any("存在しません" in error for error in validate_personal_hierarchy(rows)))

    def test_old_rows_default_to_level_one(self):
        rows = [{"個人タスクID": "P1", "タスク名": "旧データ"}]
        normalize_personal_hierarchy(rows)
        self.assertEqual(rows[0]["Level"], 1)
        self.assertEqual(rows[0]["親タスクID"], "")

    def test_level_is_derived_from_parent_id(self):
        rows = [
            {"個人タスクID": "P1", "Level": 2, "親タスクID": ""},
            {"個人タスクID": "P2", "Level": 1, "親タスクID": "P1"},
        ]
        normalize_personal_hierarchy(rows)
        self.assertEqual([row["Level"] for row in rows], [1, 2])


if __name__ == "__main__":
    unittest.main()
