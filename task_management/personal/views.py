from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any

from task_management.common.indexes import monday_iso


def as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value in (None, ""):
        return None
    try:
        return date.fromisoformat(str(value).replace("/", "-"))
    except ValueError:
        return None


def is_closed(row: dict[str, Any]) -> bool:
    return str(row.get("状態") or "") in {"完了", "中止"}


def leaf_tasks(tasks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    parent_ids = {str(row.get("親タスクID") or "").strip() for row in tasks if str(row.get("親タスクID") or "").strip()}
    return [row for row in tasks if str(row.get("個人タスクID") or "").strip() not in parent_ids]


def this_week_rows(tasks: list[dict[str, Any]], week_start: date, today: date) -> list[dict[str, Any]]:
    end = week_start + timedelta(days=6)
    output = []
    for row in leaf_tasks(tasks):
        if is_closed(row):
            continue
        due, planned = as_date(row.get("期限")), as_date(row.get("予定開始日"))
        status = str(row.get("状態") or "")
        if not ((due and due <= end) or (planned and week_start <= planned <= end) or status in {"進行中", "待ち"}):
            continue
        if status == "待ち":
            group = "C. 待ち"
        elif due and due < today:
            group = "A. 今週必須"
        elif planned and planned < week_start:
            group = "D. 持ち越し"
        elif row.get("優先度") == "高" or (due and week_start <= due <= end):
            group = "A. 今週必須"
        else:
            group = "B. 今週実施予定"
        output.append({"分類": group, "個人タスクID": row.get("個人タスクID"), "Level": row.get("Level"), "親タスクID": row.get("親タスクID"), "テーマ": row.get("テーマ名"), "タスク名": row.get("タスク名"), "次アクション": row.get("次アクション"), "タスク種別": row.get("タスク種別"), "優先度": row.get("優先度"), "状態": row.get("状態"), "期限": row.get("期限"), "見積時間": row.get("見積時間"), "待ち先": row.get("待ち先")})
    return output


def today_rows(tasks: list[dict[str, Any]], today: date) -> list[dict[str, Any]]:
    output = []
    for row in leaf_tasks(tasks):
        due, planned = as_date(row.get("期限")), as_date(row.get("予定開始日"))
        status, priority = str(row.get("状態") or ""), str(row.get("優先度") or "")
        if is_closed(row) or not (planned == today or (due and due <= today) or priority == "高" or status == "進行中"):
            continue
        rank = 1 if due and due < today else 2 if priority == "高" else 3
        output.append({"優先順位": rank, "個人タスクID": row.get("個人タスクID"), "Level": row.get("Level"), "親タスクID": row.get("親タスクID"), "タスク名": row.get("タスク名"), "次アクション": row.get("次アクション"), "テーマ": row.get("テーマ名"), "期限": row.get("期限"), "見積時間": row.get("見積時間"), "状態": row.get("状態")})
    return output


def schedule_rows(tasks: list[dict[str, Any]], week_start: date) -> list[dict[str, Any]]:
    labels = {0: "今週", 1: "来週", 2: "2週間後", 3: "3週間後"}
    output = []
    for row in leaf_tasks(tasks):
        planned = as_date(row.get("予定開始日"))
        if is_closed(row) or planned is None:
            continue
        planned_monday = date.fromisoformat(monday_iso(planned))
        weeks = (planned_monday - week_start).days // 7
        output.append({"区分": labels.get(weeks, "それ以降"), "個人タスクID": row.get("個人タスクID"), "Level": row.get("Level"), "親タスクID": row.get("親タスクID"), "テーマ": row.get("テーマ名"), "タスク名": row.get("タスク名"), "予定開始日": row.get("予定開始日"), "期限": row.get("期限"), "見積時間": row.get("見積時間"), "状態": row.get("状態"), "次アクション": row.get("次アクション")})
    return output
