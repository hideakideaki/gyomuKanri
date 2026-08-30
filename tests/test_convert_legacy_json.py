import base64
import json
import tempfile
import unittest
from pathlib import Path

from tools.convert_legacy_json import ConversionError, convert, write_vba_compatible_json


def encoded(text: str) -> str:
    return base64.b64encode(text.encode("utf-8")).decode("ascii")


class LegacyJsonConversionTests(unittest.TestCase):
    def sample(self):
        cells = []
        values = {
            "ID_L0": "L0_00001",
            "L0": "テーマA",
            "ID_L1": "L1_00001",
            "L1": "親タスク",
            "ID_L2": "L2_00001",
            "L2": "子タスク",
            "担当者": "田中",
            "2026/8/24週": "進行中",
        }
        for header, value in values.items():
            cells.append(
                {
                    "record": 1,
                    "header": encoded(header),
                    "progress": header.endswith("週"),
                    "type": "string",
                    "value": encoded(value),
                    "runs": "",
                }
            )
        return {
            "schemaVersion": 1,
            "encoding": "base64-utf8",
            "sourceWorkbook": encoded("旧管理.xlsm"),
            "exportedAt": "2026-08-24T10:00:00",
            "cells": cells,
        }

    def test_hierarchy_and_weekly_progress_are_converted(self):
        result, warnings = convert(self.sample())
        self.assertEqual(result["format"], "HierarchicalTaskManager-2")
        self.assertEqual(len(result["sheets"][0]["records"]), 1)
        self.assertEqual(len(result["sheets"][1]["records"]), 2)
        self.assertEqual(len(result["sheets"][2]["records"]), 1)
        self.assertEqual(warnings, [])

    def test_output_is_valid_json_and_records_stay_on_one_line(self):
        result, _ = convert(self.sample())
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "converted.json"
            write_vba_compatible_json(result, path)
            loaded = json.loads(path.read_text(encoding="utf-8-sig"))
            self.assertEqual(loaded["addressing"], "key-and-header")
            record_lines = [line for line in path.read_text(encoding="utf-8-sig").splitlines() if '"keyB64"' in line]
            self.assertTrue(record_lines)
            self.assertTrue(all('"fields"' in line for line in record_lines))

    def test_rejects_unsupported_schema(self):
        source = self.sample()
        source["schemaVersion"] = 2
        with self.assertRaises(ConversionError):
            convert(source)


if __name__ == "__main__":
    unittest.main()
