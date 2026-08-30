# VBA→Python対応表

## Pythonへ移行・ボタン切替済み

| ブック | 旧VBA | Python | ボタン |
|---|---|---|---|
| 共通 | `ExportWorkbookJson` | `task_management.common.service::export_workbook` | Pythonへ切替 |
| 共通 | `ImportWorkbookJson` | `task_management.common.service::import_workbook` | Pythonへ切替 |
| 共通 | `BackupWorkbook` | `task_management.common.backup::create_backup` | Pythonへ切替 |
| チーム | `AssignTaskIDs` / `AssignIssueIDs` / `AssignDecisionIDs` | `task_management.common.operations::assign_ids` | 専用ボタンがある場合のみ切替 |
| チーム | `ValidateTaskHierarchy` | `task_management.team.task::validate_hierarchy` | Pythonへ切替 |
| チーム | `SyncProgressLog` | `task_management.common.operations::sync_progress` | Pythonへ切替 |
| 個人 | `AssignPersonalTaskIDs` / `AssignManagementIDs` | `task_management.common.operations::assign_ids` | Pythonへ切替 |
| 個人 | ID重複検証 | `task_management.common.indexes::build_id_index` | 上記処理に統合 |
| 個人 | `SyncProgressLog` | `task_management.common.operations::sync_progress` | Pythonへ切替 |
| チーム | `RefreshThemeList` | `task_management.common.operations::refresh_views` | Pythonへ切替 |
| 個人 | `RefreshThisWeek` / `RefreshToday` / `RefreshSchedule` | `task_management.personal.views` | Pythonへ切替 |
| 個人 | `ProcessInbox` / `SyncCompletedTasks` | `task_management.common.operations` | Pythonへ切替 |
| 共通 | 週・日ガント更新 | `task_management.common.operations::refresh_gantt` | Pythonへ切替 |
| チーム | `SyncWeeklyGanttToTasks` | `task_management.common.operations::sync_gantt_dates` | Pythonへ切替 |

Pythonへ置換済みで参照されない旧VBAは、2026-08-29にブックから削除した。削除前の元ブックと各モジュールは `outputs/python_migration/latest/vba_export/20260829_154539` に退避している。

## 残存VBA

- チーム: `RebuildTaskStructure` と依存Utility、および `modPythonBridge`。
- 個人: `GenerateWeeklyReview` と依存Utility、および `modPythonBridge`。

上記2機能は設定シートのボタンから直接利用されるため残している。その他の通常運用処理はPythonへ切替済み。

## 将来案

- PyInstallerによる単一EXEまたはフォルダ配布形式の作成。
- EXE署名、ウイルス対策ソフトの誤検知評価、自動更新方式の設計。

PythonはMinicondaスクリプトとして運用する。
