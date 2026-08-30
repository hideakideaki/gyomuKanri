from __future__ import annotations

from contextlib import AbstractContextManager
import os
from pathlib import Path
from typing import Any

from .indexes import build_header_index, build_id_index, build_week_index
from .models import RichTextValue
from .rich_text import compress_runs, hex_to_ole_color, ole_color_to_hex


class ExcelClient(AbstractContextManager["ExcelClient"]):
    """pywin32のExcel COMを閉じ込める薄い操作層。"""

    def __init__(self, workbook_path: Path, visible: bool = False, read_only: bool = False, attach_open: bool = True):
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise RuntimeError("pywin32が必要です。Miniconda環境で pip install pywin32 を実行してください。") from exc
        target = str(workbook_path.resolve()).casefold()
        self._owns_excel = True
        self._owns_workbook = True
        self._excel = None
        self.workbook = None
        if attach_open and os.environ.get("TASKMGMT_ATTACH_OPEN", "1") != "0":
            try:
                running = win32com.client.GetObject(Class="Excel.Application")
                for workbook in running.Workbooks:
                    if str(workbook.FullName).casefold() == target:
                        self._excel = running
                        self.workbook = workbook
                        self._owns_excel = False
                        self._owns_workbook = False
                        break
            except Exception:
                pass
        if self.workbook is None:
            self._excel = win32com.client.DispatchEx("Excel.Application")
            self._excel.Visible = visible
            self._excel.DisplayAlerts = False
            self._excel.ScreenUpdating = False
            self._excel.EnableEvents = False
            self.workbook = self._excel.Workbooks.Open(str(workbook_path.resolve()), ReadOnly=read_only)
            try:
                self._excel.Calculation = -4135  # xlCalculationManual
                self._excel.CalculateBeforeSave = False
            except Exception:
                pass

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        try:
            if self._owns_workbook:
                self.workbook.Close(SaveChanges=False)
        finally:
            if self._owns_excel:
                self._excel.Quit()

    def sheet(self, name: str):
        return self.workbook.Worksheets(name)

    def headers(self, sheet_name: str, header_row: int = 2, allow_duplicates: bool = False) -> dict[str, int]:
        ws = self.sheet(sheet_name)
        last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        values = ws.Range(ws.Cells(header_row, 1), ws.Cells(header_row, last_col)).Value[0]
        if not allow_duplicates:
            return build_header_index(values)
        result: dict[str, int] = {}
        for column, value in enumerate(values, start=1):
            header = str(value).strip() if value is not None else ""
            if header and header not in result:
                result[header] = column
        return result

    def id_index(self, sheet_name: str, id_header: str, header_row: int = 2, data_row: int = 3) -> dict[str, int]:
        ws = self.sheet(sheet_name)
        column = self.headers(sheet_name, header_row)[id_header]
        last_row = ws.Cells(ws.Rows.Count, column).End(-4162).Row  # xlUp
        if last_row < data_row:
            return {}
        values = ws.Range(ws.Cells(data_row, column), ws.Cells(last_row, column)).Value
        return build_id_index(((data_row + offset, row[0]) for offset, row in enumerate(values)), id_header)

    def week_index(self, sheet_name: str, header_row: int = 2) -> dict[str, int]:
        ws = self.sheet(sheet_name)
        last_col = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        values = ws.Range(ws.Cells(header_row, 1), ws.Cells(header_row, last_col)).Value[0]
        return build_week_index(enumerate(values, start=1))

    def read_rich_text(self, cell: Any) -> RichTextValue:
        text = str(cell.Value or "")
        styles = []
        for position in range(1, len(text) + 1):
            font = cell.GetCharacters(position, 1).Font
            underline = font.Underline not in (None, False, 0, -4142)  # xlUnderlineStyleNone
            styles.append((ole_color_to_hex(font.Color), bool(font.Bold), bool(font.Italic), underline))
        return RichTextValue(text=text, runs=compress_runs(styles))

    def write_rich_text(self, cell: Any, value: RichTextValue) -> None:
        value.validate()
        cell.Value = value.text
        cell.Font.Color = hex_to_ole_color("#000000")
        cell.Font.Bold = False
        cell.Font.Italic = False
        cell.Font.Underline = -4142  # xlUnderlineStyleNone
        for run in value.runs:
            font = cell.GetCharacters(run.start, run.length).Font
            font.Color = hex_to_ole_color(run.font_color)
            font.Bold = run.bold
            font.Italic = run.italic
            font.Underline = 2 if run.underline else -4142  # xlUnderlineStyleSingle / None

    def save(self) -> None:
        self.workbook.Save()

    def read_table(self, sheet_name: str, header_row: int = 2, data_row: int = 3) -> tuple[list[str], list[dict[str, Any]]]:
        ws = self.sheet(sheet_name)
        headers = self.headers(sheet_name, header_row)
        if not headers:
            return [], []
        last_col = max(headers.values())
        last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        ordered = sorted(headers.items(), key=lambda item: item[1])
        if last_row < data_row:
            return [name for name, _ in ordered], []
        values = ws.Range(ws.Cells(data_row, 1), ws.Cells(last_row, last_col)).Value
        rows = [{name: row[column - 1] for name, column in ordered} for row in values]
        return [name for name, _ in ordered], rows
