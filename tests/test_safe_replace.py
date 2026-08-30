from pathlib import Path
from datetime import datetime

from task_management.common.confirmation import confirm_replace_import
from task_management.common.safe_replace import _equivalent


def test_replace_import_confirmation_requires_yes() -> None:
    assert confirm_replace_import(Path("book.xlsm"), Path("data.json"), lambda *_: 7) is False
    assert confirm_replace_import(Path("book.xlsm"), Path("data.json"), lambda *_: 6) is True


def test_import_verification_treats_blank_and_none_as_equal() -> None:
    assert _equivalent("", None)
    assert _equivalent(None, "")


def test_import_verification_detects_different_values() -> None:
    assert not _equivalent("未着手", "完了")


def test_import_verification_compares_legacy_datetime_as_wall_clock() -> None:
    assert _equivalent("2026-08-29T15:00:00Z", datetime(2026, 8, 29, 15, 0, 0))
