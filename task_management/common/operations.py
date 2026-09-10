from __future__ import annotations

from pathlib import Path
from datetime import date, datetime, time, timedelta
from time import sleep
from typing import Any

from task_management.personal.task import assign_management_ids, assign_personal_task_ids, normalize_personal_hierarchy, validate_personal_hierarchy
from task_management.personal.views import schedule_rows, this_week_rows, today_rows
from task_management.team.task import assign_task_ids, validate_hierarchy
from task_management.team.theme import build_theme_summary

from .backup import create_backup, make_backup_path
from .clock import local_today
from .excel_client import ExcelClient
from .excel_formats import apply_named_column_formats
from .indexes import monday_iso
from .progress_logic import ProgressSource, sync_progress_records
from .rich_text import hex_to_ole_color


EXCEL_BUSY_HRESULT = -2146777998  # 0x800AC472


def _set_gantt_fill(target: Any, color: int, attempts: int = 5) -> None:
    """Excelが一時的にビジーの場合だけ、ガントの塗りつぶしを短時間再試行する。"""
    for attempt in range(attempts):
        try:
            target.Interior.Color = color
            return
        except Exception as exc:
            hresult = getattr(exc, "hresult", exc.args[0] if exc.args else None)
            if hresult != EXCEL_BUSY_HRESULT or attempt == attempts - 1:
                raise
            sleep(0.1 * (attempt + 1))


CLEAR_DATA_LAYOUTS: dict[str, tuple[tuple[str, int], ...]] = {
    "personal": (
        ("01_個人タスク", 3),
        ("02_今週", 3),
        ("03_今日", 3),
        ("04_予定", 3),
        ("05_週次振り返り", 3),
        ("06_完了ログ", 3),
        ("08_週ガント", 5),
        ("09_日ガント", 5),
        ("10_インボックス", 3),
        ("11_進捗ログ", 3),
        ("12_課題・リスク", 3),
        ("13_意思決定", 3),
    ),
    "team": (
        ("01_テーマ一覧", 3),
        ("02_タスク", 3),
        ("03_週ガント", 5),
        ("04_日ガント", 6),
        ("05_週次進捗", 3),
        ("06_進捗ログ", 3),
        ("07_課題・リスク", 3),
        ("08_意思決定", 3),
    ),
}


def _settings(excel: ExcelClient, sheet_name: str) -> dict[str, Any]:
    _, rows = excel.read_table(sheet_name)
    return {str(row.get("設定キー") or "").strip(): row.get("設定値") for row in rows if row.get("設定キー") not in (None, "")}


def _backup_then_open(workbook: Path, backup_dir: Path) -> tuple[Path, ExcelClient]:
    backup = create_backup(workbook, backup_dir)
    return backup, ExcelClient(workbook)


def _write_records(
    excel: ExcelClient,
    sheet_name: str,
    records: list[dict[str, Any]],
    preserve_extra: bool = False,
) -> None:
    ws = excel.sheet(sheet_name)
    headers = excel.headers(sheet_name, allow_duplicates=True)
    last_col = max(headers.values())
    managed_headers = {header for record in records for header in record}
    id_header = next((name for name in ("個人タスクID", "タスクID", "テーマID") if name in headers), None)
    preserved: dict[str, dict[str, Any]] = {}
    if preserve_extra and id_header:
        id_column = headers[id_header]
        last_row = ws.Cells(ws.Rows.Count, id_column).End(-4162).Row
        if last_row >= 3:
            values = ws.Range(ws.Cells(3, 1), ws.Cells(last_row, last_col)).Value
            for row in values:
                record_id = str(row[id_column - 1] or "").strip()
                if record_id:
                    preserved[record_id] = {
                        header: row[column - 1]
                        for header, column in headers.items()
                        if header not in managed_headers and header != id_header
                    }
    old_last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    if old_last >= 3:
        ws.Range(ws.Cells(3, 1), ws.Cells(old_last, last_col)).ClearContents()
    for row_number, record in enumerate(records, start=3):
        output = dict(record)
        if id_header:
            output.update(preserved.get(str(record.get(id_header) or "").strip(), {}))
        for header, value in output.items():
            column = headers.get(header)
            if column is not None:
                if isinstance(value, date) and not isinstance(value, datetime):
                    value = datetime.combine(value, time(hour=9))
                ws.Cells(row_number, column).Value = value
    apply_named_column_formats(excel, sheet_name)


def _write_personal_levels(excel: ExcelClient, tasks: list[dict[str, Any]]) -> None:
    rows = excel.id_index("01_個人タスク", "個人タスクID")
    level_column = excel.headers("01_個人タスク")["Level"]
    for task in tasks:
        task_id = str(task.get("個人タスクID") or "").strip()
        if task_id in rows:
            excel.sheet("01_個人タスク").Cells(rows[task_id], level_column).Value = task.get("Level")


def refresh_views(workbook: Path, kind: str, backup_dir: Path, dry_run: bool = False, create_backup_before: bool = True) -> dict[str, Any]:
    with ExcelClient(workbook, read_only=dry_run) as excel:
        if kind == "team":
            _, tasks = excel.read_table("02_タスク")
            _, issues = excel.read_table("07_課題・リスク")
            _, current = excel.read_table("01_テーマ一覧")
            summaries = {str(row.get("テーマID") or ""): row.get("テーマ総括") for row in current}
            outputs = {"01_テーマ一覧": build_theme_summary(tasks, issues, summaries)}
        else:
            _, tasks = excel.read_table("01_個人タスク")
            normalize_personal_hierarchy(tasks)
            settings = _settings(excel, "07_設定")
            today = local_today()
            week_value = settings.get("今週の開始日")
            week_start = date.fromisoformat(monday_iso(week_value or today))
            outputs = {"02_今週": this_week_rows(tasks, week_start, today), "03_今日": today_rows(tasks, today), "04_予定": schedule_rows(tasks, week_start)}
        result = {"rows": {sheet: len(records) for sheet, records in outputs.items()}, "dry_run": dry_run}
        if dry_run:
            return result
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)
        if kind == "personal":
            _write_personal_levels(excel, tasks)
            apply_named_column_formats(excel, "01_個人タスク")
        for sheet, records in outputs.items():
            _write_records(excel, sheet, records, preserve_extra=True)
            if kind == "team" and sheet == "01_テーマ一覧" and records:
                ws = excel.sheet(sheet)
                headers = excel.headers(sheet)
                for header in ("未完了タスク数", "課題・リスク件数"):
                    column = headers.get(header)
                    if column is not None:
                        ws.Range(ws.Cells(3, column), ws.Cells(2 + len(records), column)).NumberFormat = "0"
                due_column = headers.get("最遅期限")
                if due_column is not None:
                    ws.Range(ws.Cells(3, due_column), ws.Cells(2 + len(records), due_column)).NumberFormat = "yyyy/m/d"
        excel.save()
        return result


def sync_completed(workbook: Path, backup_dir: Path, dry_run: bool = False, create_backup_before: bool = True) -> dict[str, Any]:
    with ExcelClient(workbook, read_only=dry_run) as excel:
        _, tasks = excel.read_table("01_個人タスク")
        _, completed = excel.read_table("06_完了ログ")
        existing = {str(row.get("個人タスクID") or "") for row in completed}
        fields = ["個人タスクID", "Level", "親タスクID", "関連テーマID", "関連チームタスクID", "タスク名", "タスク種別", "登録日", "完了日", "見積時間", "実績時間", "メモ"]
        additions = [{field: row.get(field) for field in fields} for row in tasks if row.get("状態") == "完了" and str(row.get("個人タスクID") or "") not in existing]
        result = {"added": len(additions), "dry_run": dry_run}
        if dry_run or not additions:
            return result
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)
        _write_records(excel, "06_完了ログ", completed + additions)
        excel.save()
        return result


def process_inbox(workbook: Path, backup_dir: Path, dry_run: bool = False, create_backup_before: bool = True) -> dict[str, Any]:
    with ExcelClient(workbook, read_only=dry_run) as excel:
        _, tasks = excel.read_table("01_個人タスク")
        _, inbox = excel.read_table("10_インボックス")
        candidates = [row for row in inbox if row.get("処理状態") == "タスク化" and not str(row.get("個人タスクID") or "").strip()]
        new_tasks = [{"個人タスクID": "", "Level": 1, "親タスクID": "", "タスク名": row.get("内容"), "メモ": row.get("メモ"), "登録日": local_today(), "優先度": "中", "状態": "未着手"} for row in candidates]
        combined = tasks + new_tasks
        assignments = assign_personal_task_ids(combined)
        assigned_by_row = {item.row: item.value for item in assignments}
        for offset, (source, task) in enumerate(zip(candidates, new_tasks), start=len(tasks) + 3):
            task_id = assigned_by_row[offset]
            task["個人タスクID"] = task_id
            source["個人タスクID"] = task_id
            source["処理状態"] = "処理済み"
            source["処理日"] = local_today()
        result = {"created": len(new_tasks), "dry_run": dry_run}
        if dry_run or not new_tasks:
            return result
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)
        _write_records(excel, "01_個人タスク", combined)
        _write_records(excel, "10_インボックス", inbox)
        excel.save()
        return result


def _task_period(row: dict[str, Any], kind: str) -> tuple[date | None, date | None]:
    from task_management.personal.views import as_date
    if kind == "team":
        return as_date(row.get("予定開始日")), as_date(row.get("予定終了日"))
    start = as_date(row.get("予定開始日"))
    end = as_date(row.get("期限")) or start
    return start, end


def _gantt_rows(tasks: list[dict[str, Any]], kind: str, weekly: bool) -> list[dict[str, Any]]:
    if kind == "team":
        selected = [row for row in tasks if row.get("タスクID") and (not weekly or int(row.get("Level") or 0) <= 2)]
        return [{"テーマ": row.get("テーマID"), "タスクID": row.get("タスクID"), "Level": row.get("Level"), "L1": row.get("L1"), "L2": row.get("L2"), "L3": row.get("L3"), "L4": row.get("L4"), "担当": row.get("担当"), "状態": row.get("状態"), "予定開始日": row.get("予定開始日"), "予定終了日": row.get("予定終了日")} for row in selected]
    return [{"テーマ": row.get("テーマ名"), "個人タスクID": row.get("個人タスクID"), "Level": row.get("Level"), "親タスクID": row.get("親タスクID"), "タスク名": row.get("タスク名"), "次アクション": row.get("次アクション"), "優先度": row.get("優先度"), "状態": row.get("状態"), "予定開始日": _task_period(row, kind)[0], "期限": _task_period(row, kind)[1]} for row in tasks if row.get("個人タスクID")]


def _gantt_color(source: dict[str, Any], kind: str, end: date, today: date) -> str:
    if source.get("状態") in {"完了", "中止"}:
        return "#A6A6A6"
    if end < today:
        return "#E06666"
    if source.get("状態") == "待ち":
        return "#F6B26B"
    if kind in {"team", "personal"}:
        level = int(source.get("Level") or 1)
        colors = {1: "#2F5597", 2: "#5B9BD5"} if kind == "personal" else {1: "#2F5597", 2: "#5B9BD5", 3: "#9DC3E6", 4: "#DDEBF7"}
        return colors.get(level, "#DDEBF7")
    return "#5B9BD5"


def _date_header_groups(column_dates: dict[int, date], attribute: str) -> list[tuple[int, int, int]]:
    """日付列を年または月の連続範囲にまとめる。"""
    groups: list[tuple[int, int, int]] = []
    for column, current in sorted(column_dates.items()):
        key = current.year if attribute == "year" else current.month
        if groups and groups[-1][2] == key and groups[-1][1] + 1 == column:
            start, _, previous_key = groups[-1]
            groups[-1] = (start, column, previous_key)
        else:
            groups.append((column, column, key))
    return groups


def _merge_personal_gantt_headers(ws: Any, left_end: int, column_dates: dict[int, date]) -> None:
    """個人ガントの左見出しを縦結合し、年・月を期間単位で横結合する。"""
    for column in range(1, left_end + 1):
        header_range = ws.Range(ws.Cells(2, column), ws.Cells(4, column))
        values = [ws.Cells(row, column).Value for row in range(2, 5)]
        header = next((value for value in values if value not in (None, "")), "")
        header_range.UnMerge()
        header_range.ClearContents()
        ws.Cells(2, column).Value = header
        header_range.Merge()
        header_range.HorizontalAlignment = -4108  # xlCenter
        header_range.VerticalAlignment = -4108
        header_range.WrapText = True

    first_date_column = min(column_dates)
    last_date_column = max(column_dates)
    for row, attribute, suffix in ((2, "year", "年"), (3, "month", "月")):
        timeline_range = ws.Range(ws.Cells(row, first_date_column), ws.Cells(row, last_date_column))
        timeline_range.UnMerge()
        timeline_range.ClearContents()
        for start, end, value in _date_header_groups(column_dates, attribute):
            group_range = ws.Range(ws.Cells(row, start), ws.Cells(row, end))
            ws.Cells(row, start).Value = f"{value}{suffix}"
            if end > start:
                group_range.Merge()
            group_range.HorizontalAlignment = -4108
            group_range.VerticalAlignment = -4108


def refresh_gantt(workbook: Path, kind: str, backup_dir: Path, dry_run: bool = False, create_backup_before: bool = True) -> dict[str, Any]:
    task_sheet = "02_タスク" if kind == "team" else "01_個人タスク"
    week_sheet = "03_週ガント" if kind == "team" else "08_週ガント"
    day_sheet = "04_日ガント" if kind == "team" else "09_日ガント"
    with ExcelClient(workbook, read_only=dry_run) as excel:
        _, tasks = excel.read_table(task_sheet)
        if kind == "personal":
            normalize_personal_hierarchy(tasks)
        week_records, day_records = _gantt_rows(tasks, kind, True), _gantt_rows(tasks, kind, False)
        result = {"weekly_rows": len(week_records), "daily_rows": len(day_records), "dry_run": dry_run}
        if dry_run:
            return result
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)
        if kind == "personal":
            _write_personal_levels(excel, tasks)
        for sheet_name, records, weekly, data_row in ((week_sheet, week_records, True, 5), (day_sheet, day_records, False, 6 if kind == "team" else 5)):
            ws = excel.sheet(sheet_name)
            headers = excel.headers(sheet_name)
            left_end = headers.get("予定終了日") or headers.get("期限")
            if left_end is None:
                raise ValueError(f"ガント左端ヘッダーを特定できません: {sheet_name}")
            old_last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
            if old_last >= data_row:
                managed_columns = sorted({headers[header] for record in records for header in record if header in headers})
                for column in managed_columns:
                    ws.Range(ws.Cells(data_row, column), ws.Cells(old_last, column)).ClearContents()
                ws.Range(ws.Cells(data_row, left_end + 1), ws.Cells(old_last, last_col)).Interior.ColorIndex = -4142
            if kind == "team":
                selected_sources = [row for row in tasks if row.get("タスクID") and (not weekly or int(row.get("Level") or 0) <= 2)]
            else:
                selected_sources = [row for row in tasks if row.get("個人タスクID")]
            timeline_start = None
            column_dates: dict[int, date] = {}
            periods: dict[str, int] = {}
            if weekly:
                periods = excel.week_index(sheet_name, header_row=4)
                column_dates = {column: date.fromisoformat(week) for week, column in periods.items()}
            if not weekly:
                year_text = str(ws.Cells(2, left_end + 1).Value or "").replace("年", "")
                month_text = str(ws.Cells(3, left_end + 1).Value or "").replace("月", "")
                day_text = str(ws.Cells(4, left_end + 1).Value or "")
                try:
                    timeline_start = date(int(float(year_text)), int(float(month_text)), int(float(day_text)))
                except ValueError:
                    timeline_start = min((_task_period(row, kind)[0] for row in selected_sources if _task_period(row, kind)[0]), default=local_today())
                column_dates = {
                    left_end + 1 + index: timeline_start + timedelta(days=index)
                    for index in range(last_col - left_end)
                }
            if kind == "personal" and column_dates:
                _merge_personal_gantt_headers(ws, left_end, column_dates)
            for offset, (record, source) in enumerate(zip(records, selected_sources), start=data_row):
                for header, value in record.items():
                    column = headers.get(header)
                    if column is not None:
                        if isinstance(value, date) and not isinstance(value, datetime):
                            value = datetime.combine(value, time(hour=9))
                        ws.Cells(offset, column).Value = value
                start, end = _task_period(source, kind)
                if not start or not end:
                    continue
                color = _gantt_color(source, kind, end, local_today())
                if weekly:
                    for week, column in periods.items():
                        week_start = date.fromisoformat(week)
                        if week_start <= end and week_start + timedelta(days=6) >= start:
                            _set_gantt_fill(ws.Cells(offset, column), hex_to_ole_color(color))
                else:
                    timeline_end = timeline_start + timedelta(days=last_col - left_end - 1)
                    visible_start = max(start, timeline_start)
                    visible_end = min(end, timeline_end)
                    if visible_start <= visible_end:
                        first_column = left_end + 1 + (visible_start - timeline_start).days
                        last_column = left_end + 1 + (visible_end - timeline_start).days
                        target = ws.Range(ws.Cells(offset, first_column), ws.Cells(offset, last_column))
                        _set_gantt_fill(target, hex_to_ole_color(color))
            apply_named_column_formats(excel, sheet_name, data_row=data_row)
        excel.save()
        return result


def sync_gantt_dates(workbook: Path, kind: str, backup_dir: Path, dry_run: bool = False) -> dict[str, Any]:
    if kind != "team":
        return {"updated": 0, "dry_run": dry_run, "note": "個人版は逆同期対象外"}
    with ExcelClient(workbook, read_only=dry_run) as excel:
        gantt, tasks = "03_週ガント", "02_タスク"
        gantt_rows = excel.id_index(gantt, "タスクID", data_row=5)
        task_rows = excel.id_index(tasks, "タスクID")
        weeks = excel.week_index(gantt, header_row=4)
        task_headers = excel.headers(tasks)
        updates = []
        ws = excel.sheet(gantt)
        for task_id, row in gantt_rows.items():
            marked = []
            for week, column in weeks.items():
                cell = ws.Cells(row, column)
                if cell.Interior.ColorIndex != -4142:
                    marked.append(date.fromisoformat(week))
            if marked and task_id in task_rows:
                updates.append((task_rows[task_id], min(marked), max(marked) + timedelta(days=6)))
        result = {"updated": len(updates), "dry_run": dry_run}
        if dry_run or not updates:
            return result
        backup = make_backup_path(workbook, backup_dir)
        excel.workbook.SaveCopyAs(str(backup.resolve()))
        result["backup"] = str(backup)
        task_ws = excel.sheet(tasks)
        for row, start, end in updates:
            task_ws.Cells(row, task_headers["予定開始日"]).Value = datetime.combine(start, time(hour=9))
            task_ws.Cells(row, task_headers["予定終了日"]).Value = datetime.combine(end, time(hour=9))
        apply_named_column_formats(excel, tasks)
        excel.save()
        return result


def assign_ids(workbook: Path, kind: str, backup_dir: Path, dry_run: bool = False, create_backup_before: bool = True) -> dict[str, Any]:
    if kind == "team":
        with ExcelClient(workbook, read_only=True) as excel:
            settings = _settings(excel, "09_設定")
            _, task_rows = excel.read_table("02_タスク")
            task_assignments = assign_task_ids(task_rows, int(settings.get("最大階層数") or 4))
            generic = []
            for sheet, header, prefix in (("07_課題・リスク", "課題ID", "ISSUE"), ("08_意思決定", "決定ID", "DEC")):
                _, rows = excel.read_table(sheet)
                generic.append((sheet, header, assign_management_ids(rows, header, prefix)))
        assignments = [("02_タスク", "タスクID", task_assignments), *generic]
    else:
        with ExcelClient(workbook, read_only=True) as excel:
            _, task_rows = excel.read_table("01_個人タスク")
            task_assignments = assign_personal_task_ids(task_rows)
            assignments = [("01_個人タスク", "個人タスクID", task_assignments)]
            for sheet, header, prefix in (("10_インボックス", "受付ID", "INBOX"), ("12_課題・リスク", "課題・リスクID", "ISSUE"), ("13_意思決定", "決定ID", "DEC")):
                _, rows = excel.read_table(sheet)
                assignments.append((sheet, header, assign_management_ids(rows, header, prefix)))
    count = sum(len(items) for _, _, items in assignments)
    result: dict[str, Any] = {"assigned": count, "dry_run": dry_run}
    if dry_run or count == 0:
        return result
    if create_backup_before:
        backup, excel = _backup_then_open(workbook, backup_dir)
        result["backup"] = str(backup)
    else:
        excel = ExcelClient(workbook)
    with excel:
        for sheet, header, items in assignments:
            column = excel.headers(sheet)[header]
            for item in items:
                excel.sheet(sheet).Cells(item.row, column).Value = item.value
        excel.save()
    return result


def validate_workbook(workbook: Path, kind: str) -> dict[str, Any]:
    errors: list[str] = []
    with ExcelClient(workbook, read_only=True) as excel:
        if kind == "team":
            settings = _settings(excel, "09_設定")
            _, rows = excel.read_table("02_タスク")
            errors.extend(validate_hierarchy(rows, int(settings.get("最大階層数") or 4)))
        else:
            _, rows = excel.read_table("01_個人タスク")
            normalize_personal_hierarchy(rows)
            errors.extend(validate_personal_hierarchy(rows, 2))
            for sheet, header in (("01_個人タスク", "個人タスクID"), ("10_インボックス", "受付ID"), ("12_課題・リスク", "課題・リスクID"), ("13_意思決定", "決定ID")):
                excel.id_index(sheet, header)
    return {"valid": not errors, "errors": errors, "error_count": len(errors)}


def clear_workbook_data(
    workbook: Path,
    kind: str,
    backup_dir: Path,
    dry_run: bool = False,
    create_backup_before: bool = True,
) -> dict[str, Any]:
    """設定・見出し・書式・VBAを残し、運用データだけを全シートから消去する。"""
    layouts = CLEAR_DATA_LAYOUTS[kind]
    result: dict[str, Any] = {
        "status": "dry-run" if dry_run else "success",
        "kind": kind,
        "target_sheets": [sheet for sheet, _ in layouts],
        "cleared_sheets": 0,
        "dry_run": dry_run,
    }
    if dry_run:
        return result
    with ExcelClient(workbook) as excel:
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            backup.parent.mkdir(parents=True, exist_ok=True)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)
        for sheet_name, data_row in layouts:
            ws = excel.sheet(sheet_name)
            last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            last_column = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
            if last_row < data_row:
                continue
            target = ws.Range(ws.Cells(data_row, 1), ws.Cells(last_row, last_column))
            target.ClearContents()
            if "ガント" in sheet_name:
                target.Interior.ColorIndex = -4142  # xlColorIndexNone
            result["cleared_sheets"] += 1
        excel.save()
    return result


def format_date_columns(workbook: Path, backup_dir: Path, dry_run: bool = False) -> dict[str, Any]:
    """全シートの日付列を、列位置ではなくヘッダー名に基づいて整形する。"""
    result: dict[str, Any] = {"formatted_columns": 0, "dry_run": dry_run}
    if dry_run:
        return result
    backup = create_backup(workbook, backup_dir)
    result["backup"] = str(backup)
    with ExcelClient(workbook) as excel:
        for ws in excel.workbook.Worksheets:
            result["formatted_columns"] += apply_named_column_formats(excel, str(ws.Name))
        excel.save()
    return result


def _progress_sources(excel: ExcelClient, sheet_name: str, id_header: str, attribute_map: dict[str, str]) -> list[ProgressSource]:
    ws = excel.sheet(sheet_name)
    headers = excel.headers(sheet_name)
    rows = excel.id_index(sheet_name, id_header)
    weeks = excel.week_index(sheet_name)
    sources: list[ProgressSource] = []
    for task_id, row_number in rows.items():
        attributes = {target: ws.Cells(row_number, headers[source]).Value for source, target in attribute_map.items()}
        for week, column in weeks.items():
            sources.append(ProgressSource(task_id, week, str(ws.Cells(row_number, column).Value or ""), attributes))
    return sources


def _progress_row_mapping(kind: str, task_headers: dict[str, int], progress_headers: dict[str, int]) -> dict[str, str]:
    if kind == "personal":
        candidates = {
            "テーマ": "テーマ名",
            "個人タスクID": "個人タスクID",
            "タスク名": "タスク名",
            "次アクション": "次アクション",
            "タスク種別": "タスク種別",
            "優先度": "優先度",
            "状態": "状態",
            "期限": "期限",
        }
    else:
        candidates = {
            "テーマ": "テーマID",
            "タスクID": "タスクID",
            "Level": "Level",
            "担当": "担当",
            "状態": "状態",
        }
        for header in progress_headers:
            if header.startswith("L") and header[1:].isdigit():
                candidates[header] = header
    return {
        target: source
        for target, source in candidates.items()
        if target in progress_headers and source in task_headers
    }


def _ordered_task_ids(tasks: list[dict[str, Any]], id_header: str) -> list[str]:
    return [
        str(task.get(id_header) or "").strip()
        for task in tasks
        if str(task.get(id_header) or "").strip()
    ]


def refresh_weekly_progress(workbook: Path, kind: str, backup_dir: Path, dry_run: bool = False, create_backup_before: bool = True) -> dict[str, Any]:
    """週次進捗の行をタスクマスターと同じID順に揃え、週別本文を保持する。"""
    task_sheet = "02_タスク" if kind == "team" else "01_個人タスク"
    progress_sheet = "05_週次進捗" if kind == "team" else "05_週次振り返り"
    id_header = "タスクID" if kind == "team" else "個人タスクID"
    with ExcelClient(workbook, read_only=dry_run) as excel:
        task_headers = excel.headers(task_sheet)
        progress_headers = excel.headers(progress_sheet)
        _, tasks = excel.read_table(task_sheet)
        task_ids = _ordered_task_ids(tasks, id_header)
        existing_rows = excel.id_index(progress_sheet, id_header)
        old_order = [task_id for task_id, _ in sorted(existing_rows.items(), key=lambda item: item[1])]
        result: dict[str, Any] = {
            "rows": len(task_ids),
            "added": len(set(task_ids) - set(old_order)),
            "removed": len(set(old_order) - set(task_ids)),
            "reordered": old_order != task_ids,
            "dry_run": dry_run,
        }
        if dry_run:
            return result
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)

        ws = excel.sheet(progress_sheet)
        week_columns = set(excel.week_index(progress_sheet).values())
        mapping = _progress_row_mapping(kind, task_headers, progress_headers)
        controlled_columns = {progress_headers[target] for target in mapping}
        last_column = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        old_last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        preserved: dict[str, dict[int, Any]] = {}
        rich_text: dict[tuple[str, int], Any] = {}
        hidden: dict[str, bool] = {}
        for task_id, row_number in existing_rows.items():
            hidden[task_id] = bool(ws.Rows(row_number).Hidden)
            preserved[task_id] = {}
            for column in range(1, last_column + 1):
                if column in controlled_columns:
                    continue
                value = ws.Cells(row_number, column).Value
                if column in week_columns:
                    if value not in (None, ""):
                        rich_text[(task_id, column)] = excel.read_rich_text(ws.Cells(row_number, column))
                elif value not in (None, ""):
                    preserved[task_id][column] = value

        if old_last_row >= 3:
            ws.Range(ws.Cells(3, 1), ws.Cells(old_last_row, last_column)).ClearContents()
            ws.Rows(f"3:{old_last_row}").Hidden = False
        task_by_id = {str(task.get(id_header) or "").strip(): task for task in tasks if str(task.get(id_header) or "").strip()}
        for row_number, task_id in enumerate(task_ids, start=3):
            if row_number > old_last_row and old_last_row >= 3:
                template = ws.Range(ws.Cells(3, 1), ws.Cells(3, last_column))
                target = ws.Range(ws.Cells(row_number, 1), ws.Cells(row_number, last_column))
                template.Copy()
                target.PasteSpecial(Paste=-4122)  # xlPasteFormats
            task = task_by_id[task_id]
            for target_header, source_header in mapping.items():
                ws.Cells(row_number, progress_headers[target_header]).Value = task.get(source_header)
            for column, value in preserved.get(task_id, {}).items():
                ws.Cells(row_number, column).Value = value
            for column in week_columns:
                value = rich_text.get((task_id, column))
                if value is not None:
                    excel.write_rich_text(ws.Cells(row_number, column), value)
            ws.Rows(row_number).Hidden = hidden.get(task_id, False)
        apply_named_column_formats(excel, progress_sheet)
        excel.save()
        return result


def sync_progress(workbook: Path, kind: str, backup_dir: Path, dry_run: bool = False, create_backup_before: bool = True) -> dict[str, Any]:
    if kind == "team":
        source_sheet, target_sheet, id_header, settings_sheet = "05_週次進捗", "06_進捗ログ", "タスクID", "09_設定"
        attribute_map = {"テーマ": "テーマID", "Level": "Level"}
    else:
        source_sheet, target_sheet, id_header, settings_sheet = "05_週次振り返り", "11_進捗ログ", "個人タスクID", "07_設定"
        attribute_map = {"テーマ": "テーマ", "次アクション": "次アクション", "状態": "状態"}
    with ExcelClient(workbook, read_only=dry_run) as excel:
        settings = _settings(excel, settings_sheet)
        sources = _progress_sources(excel, source_sheet, id_header, attribute_map)
        headers, existing = excel.read_table(target_sheet)
        result_model = sync_progress_records(sources, existing, id_header, str(settings.get("空欄同期時の削除可否") or "確認") == "許可")
        result: dict[str, Any] = {"added": result_model.added, "updated": result_model.updated, "deleted": result_model.deleted, "deduplicated": result_model.deduplicated, "dry_run": dry_run}
        if dry_run or not (result_model.added or result_model.updated or result_model.deleted or result_model.deduplicated):
            return result
        if create_backup_before:
            backup = make_backup_path(workbook, backup_dir)
            excel.workbook.SaveCopyAs(str(backup.resolve()))
            result["backup"] = str(backup)
        ws = excel.sheet(target_sheet)
        header_index = excel.headers(target_sheet)
        old_last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        if old_last >= 3:
            ws.Range(ws.Cells(3, 1), ws.Cells(old_last, max(header_index.values()))).ClearContents()
        for offset, record in enumerate(result_model.records, start=3):
            for header, value in record.items():
                column = header_index.get(header)
                if column is not None:
                    if isinstance(value, date) and not isinstance(value, datetime):
                        value = datetime.combine(value, time(hour=9))
                    ws.Cells(offset, column).Value = value
        apply_named_column_formats(excel, target_sheet)
        excel.save()
        return result


def configured_output_dir(workbook: Path, kind: str) -> Path:
    settings_sheet = "09_設定" if kind == "team" else "07_設定"
    with ExcelClient(workbook, read_only=True) as excel:
        configured = str(_settings(excel, settings_sheet).get("エクスポートフォルダパス") or "").strip()
    if not configured:
        return workbook.parent / "exports"
    output_dir = Path(configured)
    return output_dir if output_dir.is_absolute() else workbook.parent / output_dir


def refresh_all(workbook: Path, kind: str, dry_run: bool = False, create_backups: bool = True) -> dict[str, Any]:
    output_dir = configured_output_dir(workbook, kind)
    results: dict[str, Any] = {"output_dir": str(output_dir), "dry_run": dry_run}
    if not dry_run and create_backups:
        results["backup"] = str(create_backup(workbook, output_dir))
    if kind == "personal":
        results["process_inbox"] = process_inbox(workbook, output_dir, dry_run, create_backups)
    results["assign_ids"] = assign_ids(workbook, kind, output_dir, dry_run, create_backups)
    results["refresh_views"] = refresh_views(workbook, kind, output_dir, dry_run, create_backups)
    results["refresh_gantt"] = refresh_gantt(workbook, kind, output_dir, dry_run, create_backups)
    results["refresh_weekly_progress"] = refresh_weekly_progress(workbook, kind, output_dir, dry_run, create_backups)
    results["sync_progress"] = sync_progress(workbook, kind, output_dir, dry_run, create_backups)
    if kind == "personal":
        results["sync_completed"] = sync_completed(workbook, output_dir, dry_run, create_backups)
    if not dry_run:
        with ExcelClient(workbook) as excel:
            formatted_columns = 0
            for ws in excel.workbook.Worksheets:
                formatted_columns += apply_named_column_formats(excel, str(ws.Name))
            excel.save()
        results["formatted_date_columns"] = formatted_columns
    results["validation"] = validate_workbook(workbook, kind)
    return results
