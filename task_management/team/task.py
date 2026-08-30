from __future__ import annotations

from typing import Any

from task_management.common.id_logic import IdAssignment, assign_missing_ids


def assign_task_ids(rows: list[dict[str, Any]], max_level: int = 4) -> list[IdAssignment]:
    def needs_id(row: dict[str, Any]) -> bool:
        level = row.get("Level")
        return level not in (None, "")

    def prefix(row: dict[str, Any]) -> str:
        try:
            level = int(row["Level"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"Levelが不正です: {row.get('Level')}") from exc
        if level < 1 or level > max_level:
            raise ValueError(f"Levelが範囲外です: {level}")
        return f"L{level}"

    return assign_missing_ids(rows, "タスクID", needs_id, prefix)


def validate_hierarchy(rows: list[dict[str, Any]], max_level: int = 4, row_offset: int = 3) -> list[str]:
    errors: list[str] = []
    by_id: dict[str, tuple[int, dict[str, Any]]] = {}
    for offset, row in enumerate(rows):
        row_number = row_offset + offset
        task_id = str(row.get("タスクID") or "").strip()
        if not task_id:
            continue
        if task_id in by_id:
            errors.append(f"行 {row_number}: タスクID重複 {task_id}")
        else:
            by_id[task_id] = (row_number, row)
    for offset, row in enumerate(rows):
        row_number = row_offset + offset
        task_id = str(row.get("タスクID") or "").strip()
        if not task_id:
            continue
        try:
            level = int(row.get("Level"))
        except (TypeError, ValueError):
            errors.append(f"行 {row_number}: Levelが数値ではありません")
            continue
        if not 1 <= level <= max_level:
            errors.append(f"行 {row_number}: Levelが1～{max_level}の範囲外です")
            continue
        parent_id = str(row.get("親タスクID") or "").strip()
        if level == 1:
            if parent_id:
                errors.append(f"行 {row_number}: L1には親タスクを設定できません")
            continue
        if parent_id not in by_id:
            errors.append(f"行 {row_number}: 親タスクIDが存在しません（{parent_id}）")
            continue
        _, parent = by_id[parent_id]
        try:
            parent_level = int(parent.get("Level"))
        except (TypeError, ValueError):
            errors.append(f"行 {row_number}: 親タスクのLevelが不正です")
            continue
        if parent_level != level - 1:
            errors.append(f"行 {row_number}: 親LevelはL{level - 1}である必要があります")
        if str(parent.get("テーマID") or "") != str(row.get("テーマID") or ""):
            errors.append(f"行 {row_number}: 親子のテーマIDが一致しません")
    return errors
