from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from task_management.cli import DEFAULT_PERSONAL_WORKBOOK, DEFAULT_TEAM_WORKBOOK, configured_workbooks


class ConfiguredWorkbooksTests(unittest.TestCase):
    def test_default_names_are_explicit(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            self.assertEqual(
                configured_workbooks(root),
                [
                    ("personal", root / DEFAULT_PERSONAL_WORKBOOK),
                    ("team", root / DEFAULT_TEAM_WORKBOOK),
                ],
            )

    def test_names_can_be_changed_in_utf8_json(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / "workbooks.json").write_text(
                '{"personal":"自分.xlsm","team":"部署.xlsm"}', encoding="utf-8"
            )
            self.assertEqual(
                configured_workbooks(root),
                [("personal", root / "自分.xlsm"), ("team", root / "部署.xlsm")],
            )


if __name__ == "__main__":
    unittest.main()
