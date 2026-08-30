from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from .backup import make_backup_path
from .excel_client import ExcelClient
from .import_plan import create_plan
from .json_io import dump_v3, load_v3
from .models import Record, TableData, WeeklyProgress, WorkbookData, json_value
from .specs import WorkbookSpec


def _normalize_legacy_personal_dates(data: WorkbookData, spec: WorkbookSpec) -> None:
    if spec.kind != "personal":
        return
    for table in data.tables:
        if table.sheet != "01_個人タスク":
            continue
        for record in table.records:
            fields = record.fields
            if fields.get("予定開始日") in (None, ""):
                fields["予定開始日"] = fields.get("実施予定日") or fields.get("実施予定週")
            fields.pop("実施予定日", None)
            fields.pop("実施予定週", None)
            fields.setdefault("親タスクID", "")
            fields["Level"] = 2 if str(fields.get("親タスクID") or "").strip() else 1


def _collect_indexes(excel: ExcelClient, spec: WorkbookSpec) -> tuple[dict[str, dict[str, dict[str, int]]], dict[str, dict[str, int]], list[str]]:
    indexes: dict[str, dict[str, dict[str, int]]] = {}
    missing_sheets: list[str] = []
    for table in spec.tables:
        try:
            headers = excel.headers(table.sheet)
        except Exception:
            missing_sheets.append(table.sheet)
            continue
        rows = excel.id_index(table.sheet, table.key_header) if table.key_header in headers else {}
        indexes[table.sheet] = {"headers": headers, "rows": rows}
    try:
        weekly_headers = excel.headers(spec.weekly_sheet)
    except Exception:
        missing_sheets.append(spec.weekly_sheet)
        return indexes, {}, missing_sheets
    weekly_rows = excel.id_index(spec.weekly_sheet, spec.weekly_id_header) if spec.weekly_id_header in weekly_headers else {}
    indexes[spec.weekly_sheet] = {"headers": weekly_headers, "rows": weekly_rows}
    return indexes, {spec.weekly_sheet: excel.week_index(spec.weekly_sheet)}, missing_sheets


def _planned_additions(data: WorkbookData, indexes: dict[str, dict[str, dict[str, int]]], week_indexes: dict[str, dict[str, int]], spec: WorkbookSpec) -> tuple[list[tuple[str, str]], list[tuple[str, str]], list[str]]:
    rows: list[tuple[str, str]] = []
    columns: list[tuple[str, str]] = []
    weeks: list[str] = []
    for table in data.tables:
        sheet_index = indexes.get(table.sheet)
        if sheet_index is None:
            continue
        headers = sheet_index["headers"]
        row_index = sheet_index["rows"]
        ordered_headers = [table.key_header]
        for record in table.records:
            for header in record.fields:
                if header not in ordered_headers:
                    ordered_headers.append(header)
            if record.key not in row_index:
                rows.append((table.sheet, record.key))
        columns.extend((table.sheet, header) for header in ordered_headers if header not in headers)
    weekly_index = indexes.get(spec.weekly_sheet)
    if weekly_index is not None:
        if spec.weekly_id_header not in weekly_index["headers"]:
            columns.append((spec.weekly_sheet, spec.weekly_id_header))
        known_rows = weekly_index["rows"]
        known_weeks = week_indexes.get(spec.weekly_sheet, {})
        for item in data.weekly_progress:
            if item.task_id not in known_rows and (spec.weekly_sheet, item.task_id) not in rows:
                rows.append((spec.weekly_sheet, item.task_id))
            if item.week not in known_weeks and item.week not in weeks:
                weeks.append(item.week)
    return rows, columns, weeks


def _append_column(excel: ExcelClient, sheet_name: str, header: str) -> None:
    ws = excel.sheet(sheet_name)
    header_row = 2
    data_row = 3
    last_column = ws.Cells(header_row, ws.Columns.Count).End(-4159).Column
    new_column = last_column + 1
    last_row = max(data_row, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    source = ws.Range(ws.Cells(1, last_column), ws.Cells(last_row, last_column))
    target = ws.Range(ws.Cells(1, new_column), ws.Cells(last_row, new_column))
    source.Copy()
    target.PasteSpecial(Paste=-4122)  # xlPasteFormats
    target.ClearContents()
    ws.Columns(new_column).ColumnWidth = ws.Columns(last_column).ColumnWidth
    ws.Cells(header_row, new_column).Value = header
    excel._excel.CutCopyMode = False


def _append_row(excel: ExcelClient, sheet_name: str, key_header: str, key: str) -> None:
    ws = excel.sheet(sheet_name)
    headers = excel.headers(sheet_name)
    key_column = headers[key_header]
    last_column = max(headers.values())
    used_last_row = max(3, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    new_row = used_last_row + 1
    scan_range = ws.Range(ws.Cells(3, 1), ws.Cells(used_last_row, last_column))
    all_values = scan_range.Value
    all_formulas = scan_range.Formula
    for offset, (values, formulas) in enumerate(zip(all_values, all_formulas)):
        if all(value in (None, "") for value in values) and all(value in (None, "") for value in formulas):
            new_row = 3 + offset
            break
    target = ws.Range(ws.Cells(new_row, 1), ws.Cells(new_row, last_column))
    if new_row > used_last_row:
        source_row = max(3, new_row - 1)
        source = ws.Range(ws.Cells(source_row, 1), ws.Cells(source_row, last_column))
        source.Copy()
        target.PasteSpecial(Paste=-4122)  # xlPasteFormats
        ws.Rows(new_row).RowHeight = ws.Rows(source_row).RowHeight
        excel._excel.CutCopyMode = False
    target.ClearContents()
    ws.Cells(new_row, key_column).Value = key


def _excel_import_value(value: Any, number_format: str) -> Any:
    if not isinstance(value, str):
        return value
    normalized_format = str(number_format or "").lower()
    if "y" not in normalized_format or "d" not in normalized_format:
        return value
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return value
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(ZoneInfo("Asia/Tokyo")).replace(tzinfo=None)
    return parsed


def export_workbook(workbook: Path, output: Path, spec: WorkbookSpec) -> None:
    tables: list[TableData] = []
    progress: list[WeeklyProgress] = []
    with ExcelClient(workbook, read_only=True) as excel:
        for table_spec in spec.tables:
            ws = excel.sheet(table_spec.sheet)
            headers = excel.headers(table_spec.sheet)
            key_col = headers[table_spec.key_header]
            last_row = ws.Cells(ws.Rows.Count, key_col).End(-4162).Row
            last_col = max(headers.values())
            records: list[Record] = []
            if last_row >= 3:
                values = ws.Range(ws.Cells(3, 1), ws.Cells(last_row, last_col)).Value
                ordered_headers = [(name, column) for name, column in headers.items()]
                for row in values:
                    key = str(row[key_col - 1] or "").strip()
                    if not key:
                        continue
                    fields = {name: json_value(row[column - 1]) for name, column in ordered_headers if row[column - 1] is not None}
                    records.append(Record(key=key, fields=fields))
            tables.append(TableData(table_spec.sheet, table_spec.key_header, records))
        ws = excel.sheet(spec.weekly_sheet)
        id_rows = excel.id_index(spec.weekly_sheet, spec.weekly_id_header)
        weeks = excel.week_index(spec.weekly_sheet)
        for task_id, row in id_rows.items():
            for week, column in weeks.items():
                cell = ws.Cells(row, column)
                if cell.Value not in (None, ""):
                    progress.append(WeeklyProgress(task_id, week, excel.read_rich_text(cell)))
    dump_v3(WorkbookData(spec.format, datetime.now().isoformat(timespec="seconds"), tables, progress), output)


def import_workbook(
    workbook: Path,
    source: Path,
    spec: WorkbookSpec,
    backup_dir: Path,
    dry_run: bool = False,
    create_backup_before: bool = True,
) -> dict[str, object]:
    data = load_v3(source)
    if data.format != spec.format:
        raise ValueError(f"ブック種別とJSON formatが一致しません: {data.format}")
    _normalize_legacy_personal_dates(data, spec)
    with ExcelClient(workbook, read_only=dry_run) as excel:
        indexes, weeks, missing_sheets = _collect_indexes(excel, spec)
        plan = create_plan(data, indexes, weeks, spec.weekly_sheet)
        added_rows, added_columns, added_weeks = _planned_additions(data, indexes, weeks, spec)
        result: dict[str, Any] = {
            "updates": len(plan.updates),
            "added_rows": len(added_rows),
            "added_columns": len(added_columns),
            "added_weeks": len(added_weeks),
            "missing_ids": [],
            "missing_headers": [f"シートなし: {sheet}" for sheet in missing_sheets],
            "rich_text_errors": plan.rich_text_errors,
            "dry_run": dry_run,
        }
        if missing_sheets or plan.rich_text_errors or dry_run:
            return result
        # Excelで開いているxlsmをshutil.copy2すると、ExcelやDropboxの
        # ファイルロックによりPermissionErrorになることがある。
        # 接続中のExcel自身にコピーを作らせ、未保存内容も含む安全な退避にする。
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)
        for sheet_name, header in added_columns:
            _append_column(excel, sheet_name, header)
        for week in added_weeks:
            week_date = datetime.fromisoformat(week).date()
            _append_column(excel, spec.weekly_sheet, f"{week_date.year}/{week_date.month}/{week_date.day}週")
        indexes, weeks, missing_sheets = _collect_indexes(excel, spec)
        table_keys = {table.sheet: table.key_header for table in spec.tables}
        table_keys[spec.weekly_sheet] = spec.weekly_id_header
        for sheet_name, key in added_rows:
            _append_row(excel, sheet_name, table_keys[sheet_name], key)
        indexes, weeks, missing_sheets = _collect_indexes(excel, spec)
        plan = create_plan(data, indexes, weeks, spec.weekly_sheet)
        if not plan.valid:
            raise RuntimeError("行・列追加後のImport計画を確定できません: " + "; ".join(plan.missing_ids + plan.missing_headers + plan.rich_text_errors))
        result["updates"] = len(plan.updates)
        for update in plan.updates:
            cell = excel.sheet(update.sheet).Cells(update.row, update.column)
            if hasattr(update.value, "runs"):
                excel.write_rich_text(cell, update.value)
            else:
                cell.Value = _excel_import_value(update.value, cell.NumberFormat)
        excel.save()
        return result
