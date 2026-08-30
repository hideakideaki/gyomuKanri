#!/usr/bin/env python
"""旧タスク管理JSON（schemaVersion 1）を新形式へ変換する。"""

from __future__ import annotations

import argparse
import base64
import json
import sys
from collections import OrderedDict
from datetime import date, datetime
from pathlib import Path
from typing import Any


NEW_FORMAT = "HierarchicalTaskManager-2"
GENERAL_FORMAT = "G/標準"
DATE_FORMAT = "yyyy/m/d"


class ConversionError(ValueError):
    """入力JSONを安全に変換できない場合のエラー。"""


def b64_encode(text: str) -> str:
    return base64.b64encode(text.encode("utf-8")).decode("ascii")


def b64_decode(value: Any, label: str) -> str:
    if not isinstance(value, str):
        raise ConversionError(f"{label}が文字列ではありません。")
    try:
        return base64.b64decode(value, validate=True).decode("utf-8")
    except (ValueError, UnicodeDecodeError) as exc:
        raise ConversionError(f"{label}のBase64またはUTF-8が不正です。") from exc


def excel_serial(iso_date: str) -> str:
    try:
        parsed = date.fromisoformat(iso_date)
    except ValueError as exc:
        raise ConversionError(f"日付「{iso_date}」はYYYY-MM-DD形式ではありません。") from exc
    return str((parsed - date(1899, 12, 30)).days)


def new_field(header: str, value: Any, value_type: str = "S") -> dict[str, Any]:
    if value_type == "D":
        encoded_value = excel_serial(str(value))
        number_format = DATE_FORMAT
    elif value_type == "N":
        try:
            encoded_value = format(float(str(value)), ".15g")
        except ValueError as exc:
            raise ConversionError(f"列「{header}」の数値「{value}」が不正です。") from exc
        number_format = GENERAL_FORMAT
    elif value_type == "B":
        encoded_value = "1" if str(value).lower() in {"1", "true"} else "0"
        number_format = GENERAL_FORMAT
    else:
        value_type = "S"
        encoded_value = b64_encode(str(value))
        number_format = GENERAL_FORMAT
    return {
        "hB64": b64_encode(header),
        "t": value_type,
        "v": encoded_value,
        "fc": 0,
        "bg": 16777215,
        "b": False,
        "i": False,
        "nf": b64_encode(number_format),
    }


def add_value(record: OrderedDict[str, tuple[Any, str]], header: str, value: Any, value_type: str = "S") -> None:
    if value is None or str(value) == "":
        return
    current = record.get(header)
    candidate = (value, value_type)
    if current is not None and current != candidate:
        raise ConversionError(
            f"同じIDの列「{header}」に異なる値があります: 「{current[0]}」/「{value}」"
        )
    record[header] = candidate


def decode_legacy_records(source: dict[str, Any]) -> list[OrderedDict[str, dict[str, Any]]]:
    if source.get("schemaVersion") != 1:
        raise ConversionError("schemaVersion 1の旧JSONではありません。")
    if source.get("encoding") != "base64-utf8":
        raise ConversionError("encodingはbase64-utf8である必要があります。")
    cells = source.get("cells")
    if not isinstance(cells, list):
        raise ConversionError("cells配列がありません。")

    grouped: dict[int, OrderedDict[str, dict[str, Any]]] = {}
    for index, cell in enumerate(cells, start=1):
        if not isinstance(cell, dict):
            raise ConversionError(f"cells[{index}]がオブジェクトではありません。")
        record_no = cell.get("record")
        if not isinstance(record_no, int) or record_no < 1:
            raise ConversionError(f"cells[{index}].recordが不正です。")
        header = b64_decode(cell.get("header"), f"cells[{index}].header")
        value = b64_decode(cell.get("value"), f"cells[{index}].value")
        record = grouped.setdefault(record_no, OrderedDict())
        if header in record:
            raise ConversionError(f"record {record_no}に列「{header}」が重複しています。")
        record[header] = {
            "value": value,
            "type": str(cell.get("type", "string")),
            "progress": bool(cell.get("progress", False)),
            "runs": b64_decode(cell.get("runs", ""), f"cells[{index}].runs"),
        }
    return [grouped[number] for number in sorted(grouped)]


def legacy_value(record: OrderedDict[str, dict[str, Any]], header: str) -> str:
    cell = record.get(header)
    return "" if cell is None else str(cell["value"])


def legacy_type_to_new(type_name: str) -> str:
    return {"date": "D", "number": "N", "boolean": "B"}.get(type_name, "S")


def convert(source: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    legacy_records = decode_legacy_records(source)
    warnings: list[str] = []
    themes: OrderedDict[str, OrderedDict[str, tuple[Any, str]]] = OrderedDict()
    tasks: OrderedDict[str, OrderedDict[str, tuple[Any, str]]] = OrderedDict()
    weekly: OrderedDict[str, OrderedDict[str, tuple[Any, str]]] = OrderedDict()

    direct_map = {
        "状態": "状態",
        "担当者": "担当",
        "優先度": "優先度",
        "開始予定": "予定開始日",
        "終了予定": "予定終了日",
        "作業フォルダ": "作業フォルダ",
        "主要成果物": "主要成果物",
    }
    note_headers = (
        "進捗率",
        "次回アクション",
        "最終更新日",
        "先行タスクID",
        "タスク種別",
        "マイルストーン予定日",
    )

    for record_number, record in enumerate(legacy_records, start=1):
        theme_id = legacy_value(record, "ID_L0")
        theme_name = legacy_value(record, "L0")
        if theme_id:
            theme = themes.setdefault(theme_id, OrderedDict())
            add_value(theme, "テーマID", theme_id)
            add_value(theme, "テーマ名", theme_name or theme_id)

        chain: list[tuple[int, str, str]] = []
        for level in range(1, 5):
            task_id = legacy_value(record, f"ID_L{level}")
            task_name = legacy_value(record, f"L{level}")
            if task_id:
                chain.append((level, task_id, task_name))

        parent_id = ""
        for level, task_id, task_name in chain:
            task = tasks.setdefault(task_id, OrderedDict())
            add_value(task, "タスクID", task_id)
            add_value(task, "テーマID", theme_id)
            add_value(task, "親タスクID", parent_id)
            add_value(task, "Level", level, "N")
            add_value(task, f"L{level}", task_name or task_id)
            add_value(task, "タスク名", task_name or task_id)
            parent_id = task_id

        if not chain:
            warnings.append(f"record {record_number}: ID_L1～ID_L4がないためタスクを生成しませんでした。")
            continue

        deepest_level, deepest_id, _ = chain[-1]
        task = tasks[deepest_id]
        for old_header, new_header in direct_map.items():
            cell = record.get(old_header)
            if cell is not None:
                add_value(task, new_header, cell["value"], legacy_type_to_new(cell["type"]))

        notes = []
        for header in note_headers:
            value = legacy_value(record, header)
            if value:
                notes.append(f"{header}: {value}")
        if notes:
            existing_note = task.get("備考")
            combined = "\n".join(notes)
            if existing_note and existing_note[0] != combined:
                combined = str(existing_note[0]) + "\n" + combined
            task["備考"] = (combined, "S")

        progress_cells = [
            (header, cell) for header, cell in record.items() if cell["progress"]
        ]
        if progress_cells:
            progress = weekly.setdefault(deepest_id, OrderedDict())
            add_value(progress, "タスクID", deepest_id)
            add_value(progress, "テーマ", theme_name or theme_id)
            add_value(progress, "Level", deepest_level, "N")
            for level, _, task_name in chain:
                add_value(progress, f"L{level}", task_name)
            assignee = legacy_value(record, "担当者")
            state = legacy_value(record, "状態")
            add_value(progress, "担当", assignee)
            add_value(progress, "状態", state)
            for header, cell in progress_cells:
                add_value(progress, header, cell["value"], legacy_type_to_new(cell["type"]))
                if cell["runs"]:
                    warnings.append(
                        f"record {record_number} / {header}: 文字単位の書式情報は新形式に移せないため省略しました。"
                    )

    if not tasks:
        raise ConversionError("変換可能なタスクID（ID_L1～ID_L4）がありません。")

    def make_sheet(name: str, key_header: str, rows: OrderedDict[str, OrderedDict[str, tuple[Any, str]]]) -> dict[str, Any]:
        records = []
        for key, values in rows.items():
            records.append(
                {
                    "keyB64": b64_encode(key),
                    "fields": [new_field(header, value, typ) for header, (value, typ) in values.items()],
                }
            )
        return {
            "nameB64": b64_encode(name),
            "keyHeaderB64": b64_encode(key_header),
            "records": records,
        }

    sheets = [
        make_sheet("01_テーマ一覧", "テーマID", themes),
        make_sheet("02_タスク", "タスクID", tasks),
    ]
    if weekly:
        sheets.append(make_sheet("05_週次進捗", "タスクID", weekly))

    result = {
        "format": NEW_FORMAT,
        "addressing": "key-and-header",
        "created": datetime.now().astimezone().isoformat(timespec="seconds"),
        "convertedFrom": {
            "schemaVersion": 1,
            "sourceWorkbookB64": source.get("sourceWorkbook", ""),
            "exportedAt": source.get("exportedAt", ""),
        },
        "sheets": sheets,
    }
    return result, warnings


def write_vba_compatible_json(data: dict[str, Any], output_path: Path) -> None:
    """VBA取込処理が1行ずつ読める配置で、正しいJSONを書き出す。"""
    lines = [
        json.dumps({key: data[key] for key in ("format", "addressing", "created", "convertedFrom")}, ensure_ascii=False)[:-1]
        + ', "sheets":['
    ]
    for sheet_index, sheet in enumerate(data["sheets"]):
        sheet_head = {
            "nameB64": sheet["nameB64"],
            "keyHeaderB64": sheet["keyHeaderB64"],
        }
        lines.append(json.dumps(sheet_head, ensure_ascii=False)[:-1] + ', "records":[')
        for record_index, record in enumerate(sheet["records"]):
            suffix = "," if record_index < len(sheet["records"]) - 1 else ""
            lines.append(json.dumps(record, ensure_ascii=False, separators=(",", ":")) + suffix)
        sheet_suffix = "," if sheet_index < len(data["sheets"]) - 1 else ""
        lines.append("]}" + sheet_suffix)
    lines.append("]}")
    with output_path.open("w", encoding="utf-8-sig", newline="") as stream:
        stream.write("\n".join(lines) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="旧タスク管理JSON（schemaVersion 1）を新ツール用JSONへ変換します。"
    )
    parser.add_argument("input", type=Path, help="旧JSONファイル")
    parser.add_argument("-o", "--output", type=Path, help="出力先。省略時は入力名_converted_v2.json")
    parser.add_argument("--overwrite", action="store_true", help="出力ファイルがある場合に上書きする")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()
    output_path = (
        args.output.resolve()
        if args.output
        else input_path.with_name(input_path.stem + "_converted_v2.json")
    )
    try:
        if not input_path.is_file():
            raise ConversionError(f"入力ファイルがありません: {input_path}")
        if output_path.exists() and not args.overwrite:
            raise ConversionError(f"出力ファイルが既にあります。--overwriteを付けてください: {output_path}")
        with input_path.open("r", encoding="utf-8-sig") as stream:
            source = json.load(stream)
        if not isinstance(source, dict):
            raise ConversionError("JSONのルートがオブジェクトではありません。")
        converted, warnings = convert(source)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        write_vba_compatible_json(converted, output_path)
        with output_path.open("r", encoding="utf-8-sig") as stream:
            verified = json.load(stream)
        if verified.get("format") != NEW_FORMAT:
            raise ConversionError("出力後の検証に失敗しました。")
    except (ConversionError, json.JSONDecodeError, OSError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1

    counts = {
        b64_decode(sheet["nameB64"], "sheet.nameB64"): len(sheet["records"])
        for sheet in converted["sheets"]
    }
    print("変換が完了しました。")
    print(f"入力: {input_path}")
    print(f"出力: {output_path}")
    for sheet_name, count in counts.items():
        print(f"  {sheet_name}: {count}件")
    if warnings:
        print(f"警告: {len(warnings)}件")
        for warning in warnings:
            print(f"  - {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
