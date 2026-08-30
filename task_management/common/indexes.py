from __future__ import annotations

from datetime import date, datetime, timedelta
import re
from typing import Iterable, Sequence


def normalize_header(value: object) -> str:
    return str(value).strip() if value is not None else ""


def build_header_index(headers: Sequence[object], start_column: int = 1) -> dict[str, int]:
    result: dict[str, int] = {}
    for offset, raw in enumerate(headers):
        header = normalize_header(raw)
        if not header:
            continue
        if header in result:
            raise ValueError(f"ヘッダーが重複しています: {header}")
        result[header] = start_column + offset
    return result


def build_id_index(values: Iterable[tuple[int, object]], label: str) -> dict[str, int]:
    result: dict[str, int] = {}
    for row, raw in values:
        key = normalize_header(raw)
        if not key:
            continue
        if key in result:
            raise ValueError(f"{label}が重複しています: {key} (行 {result[key]}, {row})")
        result[key] = row
    return result


def monday_iso(value: object) -> str:
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            value = (value.replace(tzinfo=None) + timedelta(hours=9)).date()
        else:
            value = value.date()
    if isinstance(value, date):
        return (value - timedelta(days=value.weekday())).isoformat()
    text = str(value).strip().replace("週", "")
    match = re.search(r"(\d{4})[/.\-](\d{1,2})[/.\-](\d{1,2})", text)
    if match:
        parsed = date(*(int(part) for part in match.groups()))
    else:
        parsed = date.fromisoformat(text)
    return (parsed - timedelta(days=parsed.weekday())).isoformat()


def build_week_index(values: Iterable[tuple[int, object]]) -> dict[str, int]:
    result: dict[str, int] = {}
    for column, raw in values:
        if raw in (None, ""):
            continue
        try:
            week = monday_iso(raw)
        except (TypeError, ValueError):
            continue
        if week in result:
            raise ValueError(f"週ヘッダーが重複しています: {week}")
        result[week] = column
    return result
