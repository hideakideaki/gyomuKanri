from __future__ import annotations

from typing import Any


DATE_ONLY_HEADERS = frozenset({
    "登録日",
    "予定開始日",
    "予定終了日",
    "期限",
    "完了日",
    "待ち期限",
    "処理日",
    "発生日",
    "対応期限",
    "決定日",
    "見直し日",
    "週",
    "最遅期限",
})
DATETIME_HEADERS = frozenset({"更新日時"})


def number_format_for_header(header: str) -> str | None:
    if header in DATE_ONLY_HEADERS:
        return "yyyy/m/d"
    if header in DATETIME_HEADERS:
        return "yyyy/m/d h:mm"
    return None


def apply_named_column_formats(
    excel: Any,
    sheet_name: str,
    *,
    data_row: int = 3,
    minimum_last_row: int = 502,
) -> int:
    """ヘッダー名を基準に日付列の表示形式を設定する。列番号には依存しない。"""
    ws = excel.sheet(sheet_name)
    headers = excel.headers(sheet_name, allow_duplicates=True)
    used_last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    last_row = max(data_row, minimum_last_row, used_last_row)
    formatted = 0
    for header, column in headers.items():
        number_format = number_format_for_header(header)
        if number_format is None:
            continue
        ws.Range(ws.Cells(data_row, column), ws.Cells(last_row, column)).NumberFormat = number_format
        formatted += 1
    return formatted
