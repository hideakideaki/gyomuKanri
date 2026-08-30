from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TableSpec:
    sheet: str
    key_header: str


@dataclass(frozen=True)
class WorkbookSpec:
    kind: str
    format: str
    tables: tuple[TableSpec, ...]
    weekly_sheet: str
    weekly_id_header: str


TEAM = WorkbookSpec(
    kind="team",
    format="HierarchicalTaskManager-3",
    tables=(
        TableSpec("01_テーマ一覧", "テーマID"),
        TableSpec("02_タスク", "タスクID"),
        TableSpec("06_進捗ログ", "ログID"),
        TableSpec("07_課題・リスク", "課題ID"),
        TableSpec("08_意思決定", "決定ID"),
        TableSpec("09_設定", "設定キー"),
    ),
    weekly_sheet="05_週次進捗",
    weekly_id_header="タスクID",
)

PERSONAL = WorkbookSpec(
    kind="personal",
    format="PersonalTaskManager-Keyed-3",
    tables=(
        TableSpec("01_個人タスク", "個人タスクID"),
        TableSpec("10_インボックス", "受付ID"),
        TableSpec("11_進捗ログ", "ログID"),
        TableSpec("12_課題・リスク", "課題・リスクID"),
        TableSpec("13_意思決定", "決定ID"),
        TableSpec("07_設定", "設定キー"),
    ),
    weekly_sheet="05_週次振り返り",
    weekly_id_header="個人タスクID",
)


def get_spec(kind: str) -> WorkbookSpec:
    if kind == "team":
        return TEAM
    if kind == "personal":
        return PERSONAL
    raise ValueError(f"kindはteamまたはpersonalです: {kind}")
