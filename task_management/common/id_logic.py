from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Callable


ID_RE = re.compile(r"^(?P<prefix>.+)_(?P<number>\d+)$")


@dataclass(frozen=True)
class IdAssignment:
    row: int
    value: str


def next_numbers(existing_ids: list[str]) -> dict[str, int]:
    result: dict[str, int] = {}
    for value in existing_ids:
        match = ID_RE.fullmatch(value.strip())
        if match:
            prefix = match.group("prefix")
            result[prefix] = max(result.get(prefix, 1), int(match.group("number")) + 1)
    return result


def assign_missing_ids(
    rows: list[dict[str, Any]],
    id_header: str,
    needs_id: Callable[[dict[str, Any]], bool],
    prefix_for: Callable[[dict[str, Any]], str],
    row_offset: int = 3,
) -> list[IdAssignment]:
    existing: dict[str, int] = {}
    for offset, row in enumerate(rows):
        value = str(row.get(id_header) or "").strip()
        if value:
            if value in existing:
                raise ValueError(f"{id_header}が重複しています: {value} (行 {existing[value]}, {row_offset + offset})")
            existing[value] = row_offset + offset
    counters = next_numbers(list(existing))
    assignments: list[IdAssignment] = []
    for offset, row in enumerate(rows):
        if str(row.get(id_header) or "").strip() or not needs_id(row):
            continue
        prefix = prefix_for(row)
        number = counters.get(prefix, 1)
        candidate = f"{prefix}_{number:05d}"
        while candidate in existing:
            number += 1
            candidate = f"{prefix}_{number:05d}"
        counters[prefix] = number + 1
        existing[candidate] = row_offset + offset
        assignments.append(IdAssignment(row_offset + offset, candidate))
    return assignments
