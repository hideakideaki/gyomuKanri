from __future__ import annotations

from pathlib import Path
import shutil
from typing import Any

from .backup import create_backup
from .excel_client import ExcelClient
from .operations import refresh_all, validate_workbook


CURRENT_LAYOUT_VERSION = 1
VERSION_KEY = "構成バージョン"


def _version_number(value: Any) -> int:
    try:
        return max(0, int(float(value)))
    except (TypeError, ValueError):
        return 0


def _set_layout_version(workbook: Path, kind: str, version: int) -> int:
    settings_sheet = "09_設定" if kind == "team" else "07_設定"
    with ExcelClient(workbook, attach_open=False) as excel:
        ws = excel.sheet(settings_sheet)
        headers = excel.headers(settings_sheet)
        key_column = headers["設定キー"]
        value_column = headers["設定値"]
        last_row = ws.Cells(ws.Rows.Count, key_column).End(-4162).Row
        target_row = None
        previous = 0
        if last_row >= 3:
            values = ws.Range(ws.Cells(3, key_column), ws.Cells(last_row, key_column)).Value
            for offset, row in enumerate(values, start=3):
                if str(row[0] or "").strip() == VERSION_KEY:
                    target_row = offset
                    previous = _version_number(ws.Cells(offset, value_column).Value)
                    break
        if target_row is None:
            target_row = max(3, last_row + 1)
            ws.Cells(target_row, key_column).Value = VERSION_KEY
        ws.Cells(target_row, value_column).Value = version
        excel.save()
        return previous


def upgrade_workbook_layout(workbook: Path, kind: str, backup_dir: Path) -> dict[str, Any]:
    """ユーザーデータを保持したまま、登録済みの差分移行と全体更新を実施する。"""
    workbook = workbook.resolve()
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = create_backup(workbook, backup_dir)
    report: dict[str, Any] = {
        "status": "running",
        "kind": kind,
        "workbook": str(workbook),
        "backup": str(backup),
        "target_version": CURRENT_LAYOUT_VERSION,
    }
    try:
        previous = _set_layout_version(workbook, kind, CURRENT_LAYOUT_VERSION)
        report["previous_version"] = previous
        report["refresh"] = refresh_all(workbook, kind, create_backups=False)
        report["validation"] = validate_workbook(workbook, kind)
        if not report["validation"]["valid"]:
            raise RuntimeError("更新後の整合性検証に失敗しました。")
        report["status"] = "success"
        return report
    except Exception:
        shutil.copy2(backup, workbook)
        report["status"] = "rolled_back"
        raise
