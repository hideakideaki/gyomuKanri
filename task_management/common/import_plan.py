from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .models import WorkbookData


@dataclass(frozen=True)
class CellUpdate:
    sheet: str
    row: int
    column: int
    value: Any


@dataclass
class ImportPlan:
    updates: list[CellUpdate] = field(default_factory=list)
    missing_ids: list[str] = field(default_factory=list)
    missing_headers: list[str] = field(default_factory=list)
    rich_text_errors: list[str] = field(default_factory=list)

    @property
    def valid(self) -> bool:
        return not (self.missing_ids or self.missing_headers or self.rich_text_errors)


def create_plan(data: WorkbookData, indexes: dict[str, dict[str, dict[str, int]]], week_indexes: dict[str, dict[str, int]], weekly_sheet: str) -> ImportPlan:
    plan = ImportPlan()
    for table in data.tables:
        sheet_indexes = indexes.get(table.sheet)
        if not sheet_indexes:
            plan.missing_headers.append(f"シートなし: {table.sheet}")
            continue
        headers = sheet_indexes["headers"]
        rows = sheet_indexes["rows"]
        for record in table.records:
            row = rows.get(record.key)
            if row is None:
                plan.missing_ids.append(f"{table.sheet}/{record.key}")
                continue
            for header, value in record.fields.items():
                column = headers.get(header)
                if column is None:
                    plan.missing_headers.append(f"{table.sheet}/{header}")
                    continue
                plan.updates.append(CellUpdate(table.sheet, row, column, value))
    weekly = indexes.get(weekly_sheet, {})
    rows = weekly.get("rows", {})
    weeks = week_indexes.get(weekly_sheet, {})
    for item in data.weekly_progress:
        row = rows.get(item.task_id)
        column = weeks.get(item.week)
        if row is None:
            plan.missing_ids.append(f"{weekly_sheet}/{item.task_id}")
        if column is None:
            plan.missing_headers.append(f"{weekly_sheet}/週={item.week}")
        try:
            item.value.validate()
        except ValueError as exc:
            plan.rich_text_errors.append(str(exc))
        if row is not None and column is not None:
            plan.updates.append(CellUpdate(weekly_sheet, row, column, item.value))
    return plan
