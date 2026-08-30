import unittest

from task_management.common.rich_text import compress_runs, hex_to_ole_color, ole_color_to_hex


class RichTextTests(unittest.TestCase):
    def test_runs_are_compressed(self):
        styles = [("#000000", False, False, False)] * 3 + [("#FF0000", True, False, False)] * 3
        runs = compress_runs(styles)
        self.assertEqual([(r.start, r.length) for r in runs], [(1, 3), (4, 3)])

    def test_color_round_trip(self):
        self.assertEqual(ole_color_to_hex(hex_to_ole_color("#12ABEF")), "#12ABEF")


if __name__ == "__main__":
    unittest.main()
