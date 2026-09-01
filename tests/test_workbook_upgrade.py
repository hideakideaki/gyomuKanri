import unittest

from task_management.common.workbook_upgrade import _version_number


class WorkbookUpgradeTests(unittest.TestCase):
    def test_version_number_accepts_excel_numeric_values(self):
        self.assertEqual(_version_number(1.0), 1)
        self.assertEqual(_version_number("2"), 2)

    def test_version_number_treats_invalid_values_as_zero(self):
        self.assertEqual(_version_number(None), 0)
        self.assertEqual(_version_number(""), 0)


if __name__ == "__main__":
    unittest.main()
