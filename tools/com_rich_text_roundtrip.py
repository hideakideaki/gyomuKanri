from __future__ import annotations

import shutil
import sys
from pathlib import Path

from task_management.common.excel_client import ExcelClient
from task_management.common.json_io import canonical_payload, load_v3
from task_management.common.models import RichTextRun, RichTextValue
from task_management.common.service import export_workbook, import_workbook
from task_management.common.specs import get_spec


def main() -> int:
    kind, source_text, work_text = sys.argv[1:4]
    source, work_dir = Path(source_text), Path(work_text)
    spec = get_spec(kind)
    work_dir.mkdir(parents=True, exist_ok=True)
    workbook = work_dir / f"{kind}_rich_text.xlsm"
    before, after = work_dir / f"{kind}_rich_before.json", work_dir / f"{kind}_rich_after.json"
    shutil.copy2(source, workbook)
    with ExcelClient(workbook) as excel:
        rows = excel.id_index(spec.weekly_sheet, spec.weekly_id_header)
        weeks = excel.week_index(spec.weekly_sheet)
        if not rows or not weeks:
            raise RuntimeError("Rich Textテストに使えるタスク行または週列がありません")
        row, column = next(iter(rows.values())), next(iter(weeks.values()))
        value = RichTextValue(
            "AAA\nBBB\n日本語・123!",
            [
                RichTextRun(1, 4, "#000000"),
                RichTextRun(5, 3, "#FF0000", bold=True),
                RichTextRun(8, 9, "#000000"),
            ],
        )
        excel.write_rich_text(excel.sheet(spec.weekly_sheet).Cells(row, column), value)
        excel.save()
    export_workbook(workbook, before, spec)
    import_workbook(workbook, before, spec, work_dir / "backup")
    export_workbook(workbook, after, spec)
    if canonical_payload(load_v3(before)) != canonical_payload(load_v3(after)):
        raise AssertionError("Rich Text Round Tripが一致しません")
    print(f"{kind}: Rich Text Round Trip OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
