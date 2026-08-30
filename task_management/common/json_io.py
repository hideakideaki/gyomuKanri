from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .models import Record, RichTextRun, RichTextValue, TableData, WeeklyProgress, WorkbookData


SUPPORTED_FORMATS = {"HierarchicalTaskManager-3", "PersonalTaskManager-Keyed-3"}


def dump_v3(data: WorkbookData, path: Path) -> None:
    validate_workbook_data(data)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as stream:
        json.dump(asdict(data), stream, ensure_ascii=False, indent=2)
        stream.write("\n")


def load_v3(path: Path) -> WorkbookData:
    with path.open("r", encoding="utf-8-sig") as stream:
        raw = json.load(stream)
    data = WorkbookData(
        format=raw["format"],
        exported_at=raw["exported_at"],
        tables=[TableData(sheet=t["sheet"], key_header=t["key_header"], records=[Record(key=r["key"], fields=r["fields"]) for r in t["records"]]) for t in raw.get("tables", [])],
        weekly_progress=[WeeklyProgress(task_id=p["task_id"], week=p["week"], value=RichTextValue(text=p["value"]["text"], runs=[RichTextRun(**run) for run in p["value"].get("runs", [])])) for p in raw.get("weekly_progress", [])],
    )
    validate_workbook_data(data)
    return data


def validate_workbook_data(data: WorkbookData) -> None:
    if data.format not in SUPPORTED_FORMATS:
        raise ValueError(f"未対応のJSON formatです: {data.format}")
    for table in data.tables:
        seen: set[str] = set()
        if not table.sheet or not table.key_header:
            raise ValueError("sheet/key_headerは必須です")
        for record in table.records:
            if not record.key or record.key in seen:
                raise ValueError(f"空または重複したIDです: {table.sheet}/{record.key}")
            seen.add(record.key)
    seen_progress: set[tuple[str, str]] = set()
    for progress in data.weekly_progress:
        key = (progress.task_id, progress.week)
        if key in seen_progress:
            raise ValueError(f"週次進捗キーが重複しています: {key}")
        seen_progress.add(key)
        progress.value.validate()


def canonical_payload(data: WorkbookData) -> dict[str, Any]:
    payload = asdict(data)
    payload.pop("exported_at", None)
    payload["tables"] = sorted(payload["tables"], key=lambda x: (x["sheet"], x["key_header"]))
    for table in payload["tables"]:
        table["records"] = sorted(table["records"], key=lambda x: x["key"])
    payload["weekly_progress"] = sorted(payload["weekly_progress"], key=lambda x: (x["task_id"], x["week"]))
    return payload
