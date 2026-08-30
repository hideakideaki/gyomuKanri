from __future__ import annotations

import ctypes
import json
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .backup import create_backup
from .excel_client import ExcelClient
from .json_io import load_v3
from .models import WorkbookData, json_value
from .operations import clear_workbook_data, refresh_all, validate_workbook
from .service import _normalize_legacy_personal_dates, import_workbook
from .specs import WorkbookSpec


def assert_workbook_closed(workbook: Path) -> None:
    """ロールバックを保証するため、対象ブックを排他的に開けることを確認する。"""
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateFileW.restype = ctypes.c_void_p
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.CloseHandle.restype = ctypes.c_int
    handle = kernel32.CreateFileW(
        str(workbook.resolve()),
        0x80000000 | 0x40000000,
        0,
        None,
        3,
        0x80,
        None,
    )
    invalid_handle = ctypes.c_void_p(-1).value
    if handle == invalid_handle or handle == -1:
        raise PermissionError("対象Excelが開かれているか、他の処理が使用中です。Excelを閉じてから再実行してください。")
    kernel32.CloseHandle(handle)


def _equivalent(expected: Any, actual: Any) -> bool:
    if expected in (None, "") and actual in (None, ""):
        return True
    if isinstance(actual, datetime) and isinstance(expected, str):
        try:
            parsed = datetime.fromisoformat(expected.replace("Z", "+00:00"))
            # Excel COMが返すtzinfoはセル値の実タイムゾーンではないため、
            # 旧Exportの日時とExcelの壁時計時刻を比較する。
            return parsed.replace(tzinfo=None) == actual.replace(tzinfo=None)
        except ValueError:
            pass
    elif isinstance(actual, date) and isinstance(expected, str):
        try:
            expected = date.fromisoformat(expected).isoformat()
        except ValueError:
            pass
    if isinstance(actual, (date, datetime)):
        actual = json_value(actual)
    if isinstance(expected, str) and isinstance(actual, str):
        return expected.replace("Z", "+00:00") == actual.replace("Z", "+00:00")
    return expected == actual


def verify_imported_content(workbook: Path, data: WorkbookData, spec: WorkbookSpec) -> dict[str, Any]:
    mismatches: list[str] = []
    checked_cells = 0
    checked_progress = 0
    with ExcelClient(workbook, read_only=True) as excel:
        for table in data.tables:
            headers = excel.headers(table.sheet)
            rows = excel.id_index(table.sheet, table.key_header)
            for record in table.records:
                row = rows.get(record.key)
                if row is None:
                    mismatches.append(f"{table.sheet}/{record.key}: 行なし")
                    continue
                for header, expected in record.fields.items():
                    column = headers.get(header)
                    if column is None:
                        mismatches.append(f"{table.sheet}/{header}: 列なし")
                        continue
                    actual = excel.sheet(table.sheet).Cells(row, column).Value
                    checked_cells += 1
                    if not _equivalent(expected, actual):
                        mismatches.append(
                            f"{table.sheet}/{record.key}/{header}: 値不一致 "
                            f"(JSON={expected!r}, Excel={actual!r})"
                        )
        rows = excel.id_index(spec.weekly_sheet, spec.weekly_id_header)
        weeks = excel.week_index(spec.weekly_sheet)
        ws = excel.sheet(spec.weekly_sheet)
        for item in data.weekly_progress:
            row = rows.get(item.task_id)
            column = weeks.get(item.week)
            if row is None or column is None:
                mismatches.append(f"{spec.weekly_sheet}/{item.task_id}/{item.week}: セルなし")
                continue
            actual = excel.read_rich_text(ws.Cells(row, column))
            checked_progress += 1
            if actual.text != item.value.text or actual.runs != item.value.runs:
                mismatches.append(f"{spec.weekly_sheet}/{item.task_id}/{item.week}: 本文または部分書式不一致")
    return {
        "valid": not mismatches,
        "checked_cells": checked_cells,
        "checked_progress_cells": checked_progress,
        "mismatches": mismatches,
    }


def replace_import_preflight(workbook: Path, source: Path, spec: WorkbookSpec) -> dict[str, Any]:
    assert_workbook_closed(workbook)
    result = import_workbook(workbook, source, spec, workbook.parent, dry_run=True)
    result["valid"] = not (result["missing_headers"] or result["rich_text_errors"])
    return result


def safe_replace_import(workbook: Path, source: Path, spec: WorkbookSpec, backup_dir: Path) -> dict[str, Any]:
    assert_workbook_closed(workbook)
    data = load_v3(source)
    if data.format != spec.format:
        raise ValueError(f"ブック種別とJSON formatが一致しません: {data.format}")
    _normalize_legacy_personal_dates(data, spec)
    preflight = replace_import_preflight(workbook, source, spec)
    if not preflight["valid"]:
        return {"status": "preflight_failed", "preflight": preflight}

    backup = create_backup(workbook, backup_dir)
    report: dict[str, Any] = {
        "status": "running",
        "workbook": str(workbook),
        "source": str(source),
        "backup": str(backup),
        "preflight": preflight,
    }
    try:
        report["clear"] = clear_workbook_data(workbook, spec.kind, backup_dir, create_backup_before=False)
        report["import"] = import_workbook(workbook, source, spec, backup_dir, create_backup_before=False)
        verification = verify_imported_content(workbook, data, spec)
        report["import_verification"] = verification
        if not verification["valid"]:
            raise RuntimeError("Import後照合に失敗しました: " + "; ".join(verification["mismatches"][:10]))
        report["refresh"] = refresh_all(workbook, spec.kind, create_backups=False)
        validation = validate_workbook(workbook, spec.kind)
        report["validation"] = validation
        if not validation["valid"]:
            raise RuntimeError("整合性検証に失敗しました: " + "; ".join(validation["errors"][:10]))
        report["status"] = "success"
    except Exception as exc:
        report["status"] = "rolled_back"
        report["error"] = str(exc)
        try:
            shutil.copy2(backup, workbook)
            report["restored"] = True
        except Exception as restore_exc:
            report["restored"] = False
            report["restore_error"] = str(restore_exc)
    report_path = backup_dir / f"{workbook.stem}_replace_import_report_{datetime.now():%Y%m%d_%H%M%S}.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    report["report"] = str(report_path)
    return report
