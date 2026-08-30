from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Union
from zoneinfo import ZoneInfo


JsonScalar = Union[None, str, int, float, bool]


@dataclass(frozen=True)
class RichTextRun:
    start: int
    length: int
    font_color: str
    bold: bool = False
    italic: bool = False
    underline: bool = False

    def validate(self, text: str) -> None:
        if self.start < 1 or self.length < 1 or self.start + self.length - 1 > len(text):
            raise ValueError(f"Rich Text範囲が本文長と矛盾します: start={self.start}, length={self.length}, text_length={len(text)}")


@dataclass
class RichTextValue:
    text: str
    runs: list[RichTextRun] = field(default_factory=list)

    def validate(self) -> None:
        for run in self.runs:
            run.validate(self.text)


@dataclass
class Record:
    key: str
    fields: dict[str, Any]


@dataclass
class TableData:
    sheet: str
    key_header: str
    records: list[Record]


@dataclass
class WeeklyProgress:
    task_id: str
    week: str
    value: RichTextValue


@dataclass
class WorkbookData:
    format: str
    exported_at: str
    tables: list[TableData]
    weekly_progress: list[WeeklyProgress]


def json_value(value: Any) -> JsonScalar:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            value = value.astimezone(ZoneInfo("Asia/Tokyo"))
        if value.hour == value.minute == value.second == value.microsecond == 0:
            return value.date().isoformat()
        return value.isoformat(timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    raise TypeError(f"JSONへ変換できない型です: {type(value).__name__}")
