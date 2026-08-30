from __future__ import annotations

import argparse
import json
from pathlib import Path

from .common.migration import migrate_file
from .common.confirmation import confirm_clear_data
from .common.operations import assign_ids, clear_workbook_data, configured_output_dir, format_date_columns, process_inbox, refresh_all, refresh_gantt, refresh_views, sync_completed, sync_gantt_dates, sync_progress, validate_workbook
from .common.backup import create_backup
from .common.service import export_workbook, import_workbook
from .common.specs import get_spec


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description="タスク管理Excel Python連携")
    commands = root.add_subparsers(dest="command", required=True)
    export = commands.add_parser("export-json")
    export.add_argument("kind", choices=("team", "personal"))
    export.add_argument("workbook", type=Path)
    export.add_argument("output", type=Path)
    export_auto = commands.add_parser("export-auto")
    export_auto.add_argument("kind", choices=("team", "personal"))
    export_auto.add_argument("workbook", type=Path)
    imp = commands.add_parser("import-json")
    imp.add_argument("kind", choices=("team", "personal"))
    imp.add_argument("workbook", type=Path)
    imp.add_argument("source", type=Path)
    imp.add_argument("--backup-dir", type=Path, default=Path("backup"))
    imp.add_argument("--dry-run", action="store_true")
    migration = commands.add_parser("migrate-v2")
    migration.add_argument("source", type=Path)
    migration.add_argument("output", type=Path)
    for command_name in ("assign-ids", "sync-progress", "refresh-views", "refresh-gantt", "sync-gantt-dates"):
        command = commands.add_parser(command_name)
        command.add_argument("kind", choices=("team", "personal"))
        command.add_argument("workbook", type=Path)
        command.add_argument("--backup-dir", type=Path, default=Path("backup"))
        command.add_argument("--dry-run", action="store_true")
    validate = commands.add_parser("validate")
    validate.add_argument("kind", choices=("team", "personal"))
    validate.add_argument("workbook", type=Path)
    backup = commands.add_parser("backup")
    backup.add_argument("workbook", type=Path)
    backup.add_argument("--backup-dir", type=Path, default=Path("backup"))
    backup_auto = commands.add_parser("backup-auto")
    backup_auto.add_argument("kind", choices=("team", "personal"))
    backup_auto.add_argument("workbook", type=Path)
    refresh = commands.add_parser("refresh-all")
    refresh.add_argument("kind", choices=("team", "personal"))
    refresh.add_argument("workbook", type=Path)
    refresh.add_argument("--dry-run", action="store_true")
    batch = commands.add_parser("batch-folder")
    batch.add_argument("operation", choices=("export-auto", "backup-auto", "validate", "refresh-all"))
    batch.add_argument("folder", type=Path)
    for command_name in ("process-inbox", "sync-completed"):
        command = commands.add_parser(command_name)
        command.add_argument("workbook", type=Path)
        command.add_argument("--backup-dir", type=Path, default=Path("backup"))
        command.add_argument("--dry-run", action="store_true")
    date_format = commands.add_parser("format-dates")
    date_format.add_argument("workbook", type=Path)
    date_format.add_argument("--backup-dir", type=Path, default=Path("backup"))
    date_format.add_argument("--dry-run", action="store_true")
    clear_data = commands.add_parser("clear-data", help="Python専用: 設定・書式を残して全運用データを消去")
    clear_data.add_argument("kind", choices=("team", "personal"))
    clear_data.add_argument("workbook", type=Path)
    clear_data.add_argument("--backup-dir", type=Path)
    clear_data.add_argument("--dry-run", action="store_true")
    return root


def main() -> int:
    args = parser().parse_args()
    if args.command == "export-json":
        export_workbook(args.workbook, args.output, get_spec(args.kind))
        print(json.dumps({"status": "success", "output": str(args.output)}, ensure_ascii=False))
    elif args.command == "export-auto":
        from datetime import datetime
        output_dir = configured_output_dir(args.workbook, args.kind)
        output_dir.mkdir(parents=True, exist_ok=True)
        output = output_dir / f"{args.kind}_{datetime.now():%Y%m%d_%H%M%S}.json"
        export_workbook(args.workbook, output, get_spec(args.kind))
        print(json.dumps({"status": "success", "output": str(output)}, ensure_ascii=False))
    elif args.command == "import-json":
        result = import_workbook(args.workbook, args.source, get_spec(args.kind), args.backup_dir, args.dry_run)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if result["missing_ids"] or result["missing_headers"] or result["rich_text_errors"]:
            return 2
    elif args.command == "migrate-v2":
        migrate_file(args.source, args.output)
        print(json.dumps({"status": "success", "output": str(args.output)}, ensure_ascii=False))
    elif args.command == "assign-ids":
        print(json.dumps(assign_ids(args.workbook, args.kind, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "sync-progress":
        print(json.dumps(sync_progress(args.workbook, args.kind, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "refresh-views":
        print(json.dumps(refresh_views(args.workbook, args.kind, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "refresh-gantt":
        print(json.dumps(refresh_gantt(args.workbook, args.kind, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "sync-gantt-dates":
        print(json.dumps(sync_gantt_dates(args.workbook, args.kind, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "process-inbox":
        print(json.dumps(process_inbox(args.workbook, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "sync-completed":
        print(json.dumps(sync_completed(args.workbook, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "format-dates":
        print(json.dumps(format_date_columns(args.workbook, args.backup_dir, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "clear-data":
        if args.dry_run:
            result = clear_workbook_data(args.workbook, args.kind, args.backup_dir or Path("backup"), True)
        elif not confirm_clear_data(args.workbook, args.kind):
            result = {"status": "cancelled", "message": "ユーザーが全データクリアを中止しました。"}
        else:
            backup_dir = args.backup_dir or configured_output_dir(args.workbook, args.kind)
            result = clear_workbook_data(args.workbook, args.kind, backup_dir)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.command == "validate":
        result = validate_workbook(args.workbook, args.kind)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if not result["valid"]:
            return 2
    elif args.command == "backup-auto":
        output = create_backup(args.workbook, configured_output_dir(args.workbook, args.kind))
        print(json.dumps({"status": "success", "backup": str(output)}, ensure_ascii=False, indent=2))
    elif args.command == "refresh-all":
        print(json.dumps(refresh_all(args.workbook, args.kind, args.dry_run), ensure_ascii=False, indent=2))
    elif args.command == "batch-folder":
        workbooks = sorted(args.folder.glob("*_Python連携版.xlsm"))
        if not workbooks:
            print(json.dumps({"status": "error", "message": "対象Excelが見つかりません。"}, ensure_ascii=False))
            return 2
        results = []
        for workbook in workbooks:
            kind = "personal" if "個人" in workbook.name else "team"
            if args.operation == "export-auto":
                from datetime import datetime
                output_dir = configured_output_dir(workbook, kind)
                output_dir.mkdir(parents=True, exist_ok=True)
                output = output_dir / f"{kind}_{datetime.now():%Y%m%d_%H%M%S}.json"
                export_workbook(workbook, output, get_spec(kind))
                result = {"output": str(output)}
            elif args.operation == "backup-auto":
                result = {"backup": str(create_backup(workbook, configured_output_dir(workbook, kind)))}
            elif args.operation == "validate":
                result = validate_workbook(workbook, kind)
            else:
                result = refresh_all(workbook, kind)
            results.append({"kind": kind, "workbook": str(workbook), "result": result})
        print(json.dumps({"status": "success", "results": results}, ensure_ascii=False, indent=2))
    else:
        output = create_backup(args.workbook, args.backup_dir)
        print(json.dumps({"status": "success", "backup": str(output)}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
