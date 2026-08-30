from pathlib import Path

from task_management.common.confirmation import clear_confirmation_message, confirm_clear_data
from task_management.common.operations import CLEAR_DATA_LAYOUTS


def test_personal_clear_preserves_settings_sheet() -> None:
    sheets = {sheet for sheet, _ in CLEAR_DATA_LAYOUTS["personal"]}
    assert "07_設定" not in sheets
    assert {"01_個人タスク", "05_週次振り返り", "08_週ガント", "09_日ガント"} <= sheets


def test_team_clear_preserves_settings_sheet() -> None:
    sheets = {sheet for sheet, _ in CLEAR_DATA_LAYOUTS["team"]}
    assert "09_設定" not in sheets


def test_confirmation_defaults_to_explicit_yes_only() -> None:
    no = lambda *_: 7
    yes = lambda *_: 6
    workbook = Path("sample.xlsm")
    assert confirm_clear_data(workbook, "personal", no) is False
    assert confirm_clear_data(workbook, "personal", yes) is True


def test_confirmation_message_explains_backup_and_preserved_content() -> None:
    message = clear_confirmation_message(Path("sample.xlsm"), "personal")
    assert "バックアップ" in message
    assert "設定" in message
    assert "マクロ" in message
