# JSON v3仕様

formatはチーム版 `HierarchicalTaskManager-3`、個人版 `PersonalTaskManager-Keyed-3`。

通常表は `sheet`、`key_header`、`record.key`、`record.fields` で表す。row、column、addressは保存しない。値はJSON本来のnull/string/number/booleanを使い、日付・日時はISO文字列とする。

週次進捗は `task_id + week` を論理キーとする。`week` は月曜日の `YYYY-MM-DD`。本文と書式は `value.text` と `value.runs` に分ける。runのstartはExcel Charactersに合わせ1始まり。連続する同一書式は1 runに圧縮する。

```json
{
  "format": "HierarchicalTaskManager-3",
  "exported_at": "2026-08-26T12:00:00",
  "tables": [
    {
      "sheet": "02_タスク",
      "key_header": "タスクID",
      "records": [
        {"key": "L3_00001", "fields": {"タスク名": "評価方式検討", "状態": "進行中"}}
      ]
    }
  ],
  "weekly_progress": [
    {
      "task_id": "L3_00001",
      "week": "2026-08-24",
      "value": {
        "text": "AAA\nBBB\nCCC",
        "runs": [
          {"start": 1, "length": 4, "font_color": "#000000", "bold": false, "italic": false, "underline": false},
          {"start": 5, "length": 3, "font_color": "#FF0000", "bold": true, "italic": false, "underline": false}
        ]
      }
    }
  ]
}
```

通常文字列にBase64は使用しない。v2移行時はBase64をUTF-8へ復号する。v2は部分文字情報を保持していないため、移行時に復元できるのは旧JSONにあるセル全体書式までである。

Import時、JSONに存在しExcelに存在しない論理IDは新規行、通常ヘッダーは新規列、週キーは新規週列として追加する。追加前に必ずバックアップを作成し、追加後にID・ヘッダー・週索引を再構築してから値とRich Textを復元する。類似名・近い行への推測書込みは行わない。対象シートなし、ID重複、ヘッダー重複、不正runは中断条件とする。初期モードはMergeで、JSONにないExcel行は保持する。
