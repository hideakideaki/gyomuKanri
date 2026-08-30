from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from .id_logic import next_numbers
from .indexes import monday_iso


@dataclass(frozen=True)
class ProgressSource:
    task_id: str
    week: str
    body: str
    attributes: dict[str, Any]


@dataclass
class ProgressResult:
    records: list[dict[str, Any]]
    added: int = 0
    updated: int = 0
    deleted: int = 0
    deduplicated: int = 0


def _updated_at(record: dict[str, Any]) -> datetime:
    value = record.get("更新日時")
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if value not in (None, ""):
        try:
            return datetime.fromisoformat(str(value).replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            pass
    return datetime.min


def sync_progress_records(
    sources: list[ProgressSource],
    existing: list[dict[str, Any]],
    task_id_header: str,
    delete_blanks: bool = False,
) -> ProgressResult:
    records: list[dict[str, Any]] = []
    index: dict[tuple[str, str], int] = {}
    deduplicated = 0
    for source_record in existing:
        record = dict(source_record)
        if all(value in (None, "") for value in record.values()):
            continue
        task_id = str(record.get(task_id_header) or "").strip()
        week_raw = record.get("週")
        if not task_id or week_raw in (None, ""):
            records.append(record)
            continue
        key = (task_id, monday_iso(week_raw))
        if key in index:
            position = index[key]
            if _updated_at(record) >= _updated_at(records[position]):
                records[position] = record
            deduplicated += 1
            continue
        index[key] = len(records)
        records.append(record)
    counters = next_numbers([str(record.get("ログID") or "") for record in existing])
    next_log = counters.get("LOG", 1)
    result = ProgressResult(records, deduplicated=deduplicated)
    delete_positions: set[int] = set()
    for source in sources:
        key = (source.task_id.strip(), monday_iso(source.week))
        position = index.get(key)
        body = source.body.strip()
        if not body:
            if delete_blanks and position is not None:
                delete_positions.add(position)
                result.deleted += 1
            continue
        values = {task_id_header: key[0], "週": date.fromisoformat(key[1]), "進捗本文": source.body, **source.attributes, "更新日時": datetime.now().replace(microsecond=0)}
        if position is None:
            values["ログID"] = f"LOG_{next_log:05d}"
            next_log += 1
            index[key] = len(records)
            records.append(values)
            result.added += 1
        else:
            records[position].update(values)
            result.updated += 1
    if delete_positions:
        result.records = [record for position, record in enumerate(records) if position not in delete_positions]
    return result
