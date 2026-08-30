from __future__ import annotations

from typing import Any

from task_management.common.clock import local_today


def build_theme_summary(tasks: list[dict[str, Any]], issues: list[dict[str, Any]], summaries: dict[str, Any]) -> list[dict[str, Any]]:
    open_counts: dict[str, int] = {}
    issue_counts: dict[str, int] = {}
    for row in tasks:
        theme = str(row.get("テーマID") or "")
        if theme and row.get("状態") not in {"完了", "中止"}:
            open_counts[theme] = open_counts.get(theme, 0) + 1
    for row in issues:
        theme = str(row.get("テーマID") or "")
        if theme and row.get("状態") not in {"完了", "中止"}:
            issue_counts[theme] = issue_counts.get(theme, 0) + 1
    output = []
    for row in tasks:
        if int(row.get("Level") or 0) != 1:
            continue
        theme = str(row.get("テーマID") or "")
        output.append({"テーマID": theme, "テーマ名": row.get("L1") or row.get("タスク名"), "責任者": row.get("担当"), "状態": row.get("状態"), "予定開始日": row.get("予定開始日"), "予定終了日": row.get("予定終了日"), "未完了タスク数": open_counts.get(theme, 0), "課題・リスク件数": issue_counts.get(theme, 0), "更新日": local_today(), "テーマ総括": summaries.get(theme)})
    return output
