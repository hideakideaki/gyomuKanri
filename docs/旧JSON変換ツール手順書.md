# 旧JSON変換ツール 手順書

## 目的

従来のタスク管理Excelが出力した `schemaVersion: 1` のJSONを、複数テーマ・階層型タスク管理Excelが読み込める `HierarchicalTaskManager-2` 形式へ変換します。

変換元JSONは変更しません。別名のJSONを新規作成します。

## 変換内容

- `ID_L0` と `L0` から「01_テーマ一覧」を生成します。
- `ID_L1`～`ID_L4`を重複排除し、「02_タスク」にID単位で展開します。
- Level、親タスクID、テーマIDは階層から自動生成します。
- L1～L4は該当Levelの列だけに設定します。
- 旧JSONの週列は「05_週次進捗」へタスクID単位で移します。
- 状態、担当者、優先度、予定日、作業フォルダ、主要成果物を新しい列名へ対応付けます。
- 新ブックに専用列がない旧項目は、値を失わないよう「備考」に項目名付きでまとめます。
- 文字単位の装飾情報 `runs` は新JSON形式に対応先がないため省略し、警告を表示します。セルの文字内容は維持します。

## 実行方法（Windows・Miniconda）

1. Anaconda PromptまたはMiniconda Promptを開きます。
2. 次のフォルダへ移動します。

```bat
cd /d "C:\Dropbox\301_CodeX\業務管理について3"
```

3. 変換を実行します。

```bat
python tools\convert_legacy_json.py "C:\Dropbox\301_CodeX\業務管理について\backup\TaskMigration_20260820_225621.json"
```

入力JSONと同じフォルダに、次のようなファイルが作成されます。

```text
TaskMigration_20260820_225621_converted_v2.json
```

保存先を指定する場合は `-o` を使います。

```bat
python tools\convert_legacy_json.py "変換元.json" -o "C:\任意のフォルダ\変換後.json"
```

同名ファイルを意図的に上書きする場合だけ `--overwrite` を追加します。

## Excelへの取込み

1. 変換後のJSONを、複数テーマ・階層型タスク管理ExcelのJSONインポートで選択します。
2. 「01_テーマ一覧」「02_タスク」「05_週次進捗」を確認します。
3. 特に、テーマID、タスクID、親タスクID、Level、週別進捗を確認します。
4. 内容確認後にExcelを別名保存します。

## 注意

- 取込み前に対象Excelのバックアップを作成してください。
- 同じタスクIDに矛盾するテーマ、親、名称などがある場合は、誤った統合を避けるため変換を停止します。
- 変換元はUTF-8のJSONを想定しています。
