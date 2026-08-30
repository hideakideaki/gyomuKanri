from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL_NS = {"p": "http://schemas.openxmlformats.org/package/2006/relationships"}
CELL_RE = re.compile(r"([A-Z]+)(\d+)")


def col_number(address: str) -> int:
    letters = CELL_RE.fullmatch(address).group(1)
    result = 0
    for char in letters:
        result = result * 26 + ord(char) - 64
    return result


def text_of(cell: ET.Element, shared: list[str]) -> object:
    cell_type = cell.get("t")
    value = cell.find("m:v", NS)
    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(".//m:t", NS))
    if value is None:
        return None
    raw = value.text or ""
    if cell_type == "s":
        return shared[int(raw)]
    if cell_type == "b":
        return raw == "1"
    if cell_type == "str":
        return raw
    try:
        return float(raw) if "." in raw else int(raw)
    except ValueError:
        return raw


def analyze(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.findall(".//m:t", NS)) for si in root.findall("m:si", NS)]

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.get("Id"): r.get("Target") for r in rels.findall("p:Relationship", REL_NS)}
        sheets = []
        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            name = sheet.get("name")
            target = rel_map[sheet.get(f"{{{NS['r']}}}id")].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            root = ET.fromstring(zf.read(target))
            cells_by_row: dict[int, dict[int, object]] = {}
            for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
                match = CELL_RE.fullmatch(cell.get("r", ""))
                if not match:
                    continue
                row = int(match.group(2))
                cells_by_row.setdefault(row, {})[col_number(cell.get("r"))] = text_of(cell, shared)
            header_row = 2 if 2 in cells_by_row else (min(cells_by_row) if cells_by_row else None)
            headers = []
            if header_row:
                headers = [str(v) for _, v in sorted(cells_by_row[header_row].items()) if v not in (None, "")]
            dim = root.find("m:dimension", NS)
            sheets.append({"name": name, "state": sheet.get("state", "visible"), "dimension": dim.get("ref") if dim is not None else None, "header_row": header_row, "headers": headers})

        defined_names = []
        for dn in workbook.findall("m:definedNames/m:definedName", NS):
            defined_names.append({"name": dn.get("name"), "localSheetId": dn.get("localSheetId"), "refers_to": dn.text})

        macro_names = set()
        control_parts = [n for n in zf.namelist() if n.startswith(("xl/drawings/", "xl/ctrlProps/")) and n.endswith(".xml")]
        macro_pattern = re.compile(r'(?:macro|fmlaMacro)="([^"]+)"')
        for part in control_parts:
            text = zf.read(part).decode("utf-8", errors="replace")
            macro_names.update(macro_pattern.findall(text))

        return {
            "path": str(path.resolve()),
            "size": path.stat().st_size,
            "sheets": sheets,
            "defined_names": defined_names,
            "assigned_macros_in_xml": sorted(macro_names),
            "has_vba_project": "xl/vbaProject.bin" in zf.namelist(),
            "vba_project_size": len(zf.read("xl/vbaProject.bin")) if "xl/vbaProject.bin" in zf.namelist() else 0,
        }


def main() -> int:
    paths = [Path(arg) for arg in sys.argv[1:]]
    print(json.dumps([analyze(path) for path in paths], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
