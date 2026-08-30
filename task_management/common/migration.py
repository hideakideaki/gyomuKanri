from __future__ import annotations

import base64
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from .json_io import dump_v3
from .models import Record, RichTextRun, RichTextValue, TableData, WeeklyProgress, WorkbookData
from .rich_text import ole_color_to_hex


def decode_b64(value: str) -> str:
    return base64.b64decode(value).decode("utf-8")


def _legacy_value(field: dict[str, Any]) -> Any:
    kind, value = field.get("t"), field.get("v")
    if kind == "S":
        return decode_b64(value)
    if kind == "B":
        return value in ("1", 1, True)
    if kind == "N":
        number = float(value)
        return int(number) if number.is_integer() else number
    if kind == "D":
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date().isoformat()
    return value


def migrate_team_v2(raw: dict[str, Any]) -> WorkbookData:
    tables: list[TableData] = []
    progress: list[WeeklyProgress] = []
    for sheet in raw.get("sheets", []):
        name = decode_b64(sheet["nameB64"])
        key_header = decode_b64(sheet["keyHeaderB64"])
        records = []
        for legacy in sheet.get("records", []):
            key = decode_b64(legacy["keyB64"])
            fields = {}
            for field in legacy.get("fields", []):
                header = decode_b64(field["hB64"])
                value = _legacy_value(field)
                if name == "05_週次進捗" and header.endswith("週") and isinstance(value, str):
                    week = header.removesuffix("週").replace("/", "-")
                    run = RichTextRun(1, len(value), ole_color_to_hex(field.get("fc", 0)), bool(field.get("b")), bool(field.get("i"))) if value else None
                    progress.append(WeeklyProgress(key, week, RichTextValue(value, [run] if run else [])))
                else:
                    fields[header] = value
            if name != "05_週次進捗":
                records.append(Record(key, fields))
        if name != "05_週次進捗":
            tables.append(TableData(name, key_header, records))
    return WorkbookData("HierarchicalTaskManager-3", datetime.now().isoformat(timespec="seconds"), tables, progress)


def migrate_personal_v2(raw: dict[str, Any]) -> WorkbookData:
    grouped: dict[tuple[str, str], dict[str, Record]] = {}
    progress: list[WeeklyProgress] = []
    for item in raw.get("records", []):
        if not item:
            continue
        sheet, key_header = decode_b64(item["s"]), decode_b64(item["kh"])
        if sheet == "05_週次振り返り":
            text = _legacy_value(item)
            if isinstance(text, str):
                run = RichTextRun(1, len(text), ole_color_to_hex(item.get("fc", 0)), bool(item.get("b")), bool(item.get("i"))) if text else None
                progress.append(WeeklyProgress(decode_b64(item["rk"]), decode_b64(item["cn"]), RichTextValue(text, [run] if run else [])))
            continue
        if key_header == "設定リスト":
            continue
        key, column = decode_b64(item["rk"]), decode_b64(item["cn"])
        record = grouped.setdefault((sheet, key_header), {}).setdefault(key, Record(key, {}))
        record.fields[column] = _legacy_value(item)
    tables = [TableData(sheet, key_header, list(records.values())) for (sheet, key_header), records in grouped.items()]
    return WorkbookData("PersonalTaskManager-Keyed-3", datetime.now().isoformat(timespec="seconds"), tables, progress)


def migrate_file(source: Path, target: Path) -> None:
    with source.open("r", encoding="utf-8-sig") as stream:
        raw = json.load(stream)
    if raw.get("format") == "HierarchicalTaskManager-2":
        data = migrate_team_v2(raw)
    elif raw.get("format") == "PersonalTaskManager-Keyed-2":
        data = migrate_personal_v2(raw)
    else:
        raise ValueError(f"未対応の旧formatです: {raw.get('format')}")
    dump_v3(data, target)
