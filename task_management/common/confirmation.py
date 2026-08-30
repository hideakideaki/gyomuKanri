from __future__ import annotations

import ctypes
from pathlib import Path
from typing import Callable


def clear_confirmation_message(workbook: Path, kind: str) -> str:
    label = "個人タスク管理" if kind == "personal" else "複数テーマ管理"
    return (
        f"{label}ブックのユーザーデータをすべて消去します。\n\n"
        f"対象:\n{workbook.resolve()}\n\n"
        "消去前にバックアップを作成します。\n"
        "シート、見出し、書式、設定、マクロは残ります。\n\n"
        "本当に実行しますか？"
    )


def confirm_clear_data(
    workbook: Path,
    kind: str,
    message_box: Callable[[int, str, str, int], int] | None = None,
) -> bool:
    """既定ボタンを「いいえ」にした警告ダイアログで最終確認する。"""
    box = message_box or ctypes.windll.user32.MessageBoxW
    flags = 0x00000004 | 0x00000030 | 0x00000100 | 0x00001000
    result = box(0, clear_confirmation_message(workbook, kind), "全データクリアの確認", flags)
    return result == 6  # IDYES


def confirm_replace_import(
    workbook: Path,
    source: Path,
    message_box: Callable[[int, str, str, int], int] | None = None,
) -> bool:
    box = message_box or ctypes.windll.user32.MessageBoxW
    message = (
        "現在の運用データを全消去し、指定JSONから復元します。\n\n"
        f"Excel:\n{workbook.resolve()}\n\n"
        f"JSON:\n{source.resolve()}\n\n"
        "処理内容:\n"
        "1. 事前検証\n2. バックアップ\n3. 全データクリア\n"
        "4. JSON Import\n5. 全体更新\n6. 整合性検証\n\n"
        "途中で失敗した場合はバックアップから自動復元します。\n"
        "本当に実行しますか？"
    )
    flags = 0x00000004 | 0x00000030 | 0x00000100 | 0x00001000
    return box(0, message, "全クリア＋Importの最終確認", flags) == 6
