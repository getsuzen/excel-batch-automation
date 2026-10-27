# Excelデータ一括統合・重複除去ツール

## 概要

同一フォーマットの複数の Excel ファイルを一括で読み込み、1つのファイルに統合する Python 製の業務自動化ツールです。

各ファイルの先頭ワークシートからデータを取得し、全列の値が一致する重複行を除外して `output/summary.xlsx` に出力します。月次データ、拠点別実績、担当者別リストなど、共通の列構成を持つ Excel ファイルの集約を想定しています。

## 主な機能

- `input` フォルダー内の `.xlsx` ファイルをファイル名順に読み込み
- 各ファイルの先頭ワークシートを統合
- 共通のヘッダーを1回だけ出力
- 全列の値が一致する重複行を除外し、最初に読み込んだ行を保持
- 全セルが空の行と、`~$` で始まる Excel 一時ファイルを除外
- 空のブック、破損ファイル、ヘッダーが一致しないファイルをスキップ
- 有効ファイル数、統合前後の行数、除外した重複行数を実行時に表示
- 出力フォルダーを自動作成し、結果を `output/summary.xlsx` に保存
- 出力ファイルにヘッダー書式、オートフィルター、先頭行の固定、列幅調整を適用

## 使用技術

- Python 3
- openpyxl 3.1.5
- Python 標準ライブラリ（`pathlib`、`zipfile`、`xml.etree.ElementTree`）
- Windows PowerShell（動作確認環境）

## ディレクトリ構成

```text
excel-batch-automation/
├─ input/
│  ├─ sales_01.xlsx
│  ├─ sales_02.xlsx
│  └─ sales_03.xlsx
├─ output/
│  └─ summary.xlsx
├─ merge_excel.py
├─ create_sample_files.py
├─ requirements.txt
└─ README.md
```

| パス | 内容 |
|---|---|
| `merge_excel.py` | Excel ファイルの統合・重複除去処理 |
| `create_sample_files.py` | 動作確認用サンプルファイルの生成 |
| `input/` | 統合対象ファイルの配置先 |
| `output/` | 処理結果の出力先 |

`input` と `output` 内の Excel ファイルは Git の管理対象外です。

## セットアップ

Windows PowerShell でプロジェクトフォルダーに移動し、仮想環境と依存パッケージを準備します。

```powershell
cd path\to\excel-batch-automation
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 実行方法

1. 統合対象の Excel ファイルを `input` フォルダー直下に配置します。
2. 次のコマンドを実行します。

```powershell
python merge_excel.py
```

処理が正常に完了すると、次のファイルが生成されます。

```text
output\summary.xlsx
```

既存の `output/summary.xlsx` がある場合は、正常実行時に上書きされます。

### サンプルデータで確認する場合

```powershell
python create_sample_files.py
python merge_excel.py
```

`create_sample_files.py` は、`input` フォルダー内の同名サンプルファイルを上書きします。

## 入力ファイル条件

- ファイル形式が `.xlsx` であること
- 対象データが各ファイルの先頭ワークシートにあること
- 先頭の有効行がヘッダーで、その次の行からデータが始まること
- すべてのファイルでヘッダーの値と列順が一致していること
- 統合対象ファイルが `input` フォルダー直下に配置されていること

全セルが空の行は処理対象外です。空のブック、読み込めないファイル、ヘッダーが一致しないファイルはスキップし、その他のファイルの処理を継続します。

処理可能なファイルがない場合、新しい出力ファイルは生成されません。以前の出力ファイルが存在する場合は、そのまま保持されます。

## 動作確認結果

同梱のサンプルファイル3件を使用して動作確認しています。サンプルデータには、重複確認用として同一内容の行が1件含まれています。

| 確認項目 | 結果 |
|---|---:|
| 対象ファイル数 | 3 |
| 正常に読み込んだファイル数 | 3 |
| 統合前のデータ行数 | 7 |
| 除外した重複行数 | 1 |
| 統合後のデータ行数 | 6 |
| 出力ファイル | `output/summary.xlsx` |
