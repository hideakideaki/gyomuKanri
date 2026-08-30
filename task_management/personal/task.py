from __future__ import annotations

from typing import Any

from task_management.common.id_logic import IdAssignment, assign_missing_ids


def assign_personal_task_ids(rows: list[dict[str, Any]]) -> list[IdAssignment]:
    return assign_missing_ids(
        rows,
        "個人タスクID",
        lambda row: bool(str(row.get("タスク名") or "").strip()),
        lambda row: "PT",
    )


def normalize_personal_hierarchy(rows: list[dict[str, Any]]) -> None:
    """親タスクIDの有無からLevelを一意に導出する。"""
    for row in rows:
        if not str(row.get("個人タスクID") or "").strip():
            continue
        row.setdefault("親タスクID", "")
        row["Level"] = 2 if str(row.get("親タスクID") or "").strip() else 1


def validate_personal_hierarchy(rows: list[dict[str, Any]], max_level: int = 2, row_offset: int = 3) -> list[str]:
    active = [row for row in rows if str(row.get("個人タスクID") or "").strip()]
    by_id = {str(row.get("個人タスクID") or "").strip(): row for row in active}
    errors: list[str] = []
    for index, row in enumerate(active):
        row_number = index + row_offset
        task_id = str(row.get("個人タスクID") or "").strip()
        parent_id = str(row.get("親タスクID") or "").strip()
        try:
            level = int(row.get("Level"))
        except (TypeError, ValueError):
            errors.append(f"行 {row_number}: Levelが数値ではありません")
            continue
        if not 1 <= level <= max_level:
            errors.append(f"行 {row_number}: Levelが1～{max_level}の範囲外です")
            continue
        if level == 1 and parent_id:
            errors.append(f"行 {row_number}: Level 1には親タスクを設定できません")
        if level > 1 and not parent_id:
            errors.append(f"行 {row_number}: Level {level}には親タスクIDが必要です")
        if parent_id == task_id:
            errors.append(f"行 {row_number}: 自分自身を親タスクにはできません")
            continue
        if parent_id and parent_id not in by_id:
            errors.append(f"行 {row_number}: 親タスクIDが存在しません（{parent_id}）")
            continue
        if parent_id:
            try:
                parent_level = int(by_id[parent_id].get("Level"))
            except (TypeError, ValueError):
                errors.append(f"行 {row_number}: 親タスクのLevelが不正です")
                continue
            if parent_level != level - 1:
                errors.append(f"行 {row_number}: 親LevelはLevel {level - 1}である必要があります")
    return errors


def assign_management_ids(rows: list[dict[str, Any]], id_header: str, default_prefix: str) -> list[IdAssignment]:
    def has_data(row: dict[str, Any]) -> bool:
        return any(value not in (None, "") for name, value in row.items() if name != id_header)

    def prefix(row: dict[str, Any]) -> str:
        return "RISK" if row.get("種別") == "リスク" else default_prefix

    return assign_missing_ids(rows, id_header, has_data, prefix)
