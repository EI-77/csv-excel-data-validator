# CSV Excel Data Validator

CSV・Excelファイルのデータ不備を検査し，Excelレポートとして出力するPythonツールです．

`input` フォルダに配置したCSV・Excelファイルを読み込み，必須列の不足，空欄セル，重複行，メールアドレス形式の不備，URL形式の不備，数値列の不正値を検出できます．

検査結果は `output` フォルダにExcelレポートとして出力されます．
レポートには，全体サマリー，ファイル別サマリー，不備の詳細，読み込みエラー，検査時の設定内容がシート別にまとめられます．

企業リスト，顧客リスト，商品リスト，店舗リスト，納品前データなどの品質チェックに利用できます．

---

## 主な機能

* CSVファイルの読み込み
* Excelファイルの読み込み
* 必須列の不足チェック
* 空欄セルのチェック
* 重複行のチェック
* メールアドレス形式のチェック
* URL形式のチェック
* 数値列の形式チェック
* ファイルごとのOK / NG / ERROR判定
* 検査結果のExcelレポート出力
* 検査時の設定内容の出力
* Excel列幅の自動調整
* Excel見出し行の装飾
* Excelフィルター設定
* Excel 1行目固定

---

## 使用技術

* Python
* pandas
* xlsxwriter
* openpyxl

---

## 想定用途

このツールは，以下のようなデータ確認作業を自動化するためのものです．

* 企業リストの納品前チェック
* 顧客リストの空欄確認
* メールアドレス一覧の形式チェック
* URL一覧の形式チェック
* 商品リストの価格列チェック
* CSV・Excelデータの重複確認
* 複数ファイルの品質確認
* 手作業で行っていたExcelチェックの自動化

---

## ディレクトリ構成

```text
csv-excel-data-validator/
├─ csv_excel_data_validator.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
│  └─ sample_companies.csv
└─ output/
   └─ validation_report_YYYYMMDD_HHMMSS.xlsx
```

---

## 実行方法

### 1. リポジトリを取得

```bash
git clone <repository_url>
```

### 2. ディレクトリへ移動

```bash
cd csv-excel-data-validator
```

### 3. 必要ライブラリをインストール

```bash
pip install -r requirements.txt
```

### 4. 検査したいファイルを配置

検査したいCSV・Excelファイルを `input` フォルダに配置します．

例：

```text
input/
├─ company_list.csv
├─ customer_list.xlsx
└─ product_list.csv
```

### 5. 設定を確認

検査条件は `config.json` で設定します．

初回実行時に `config.json` が存在しない場合は，サンプル設定ファイルが自動作成されます．

### 6. プログラムを実行

```bash
python csv_excel_data_validator.py
```

または

```bash
py csv_excel_data_validator.py
```

### 7. レポートを確認

実行後，`output` フォルダにExcelレポートが出力されます．

```text
output/
└─ validation_report_20260612_150932.xlsx
```

---

## 入力ファイルの形式

このツールは，CSV・Excelファイルが通常の表形式になっていることを前提としています．

* 1行目に列名があること
* 2行目以降にデータがあること
* Excelファイルの場合，表はA1セルから始まること
* ヘッダーより前にタイトル行・空白行がないこと
* 1つのファイルに1つの表があること

例：

| company_name | email                                             | url                   | price |
| ------------ | ------------------------------------------------- | --------------------- | ----- |
| ABC株式会社      | [info@example.com](mailto:info@example.com)       | https://example.com   | 10000 |
| XYZ株式会社      | [contact@example.com](mailto:contact@example.com) | https://example.co.jp | 20000 |

### 注意

このツールは，どんなExcelファイルでも自動で解析する万能ツールではありません．

以下のようなファイルは，正しく読み込めない場合があります．

* 表がA1セル以外から始まっているExcel
* 1行目より前にタイトルや説明文があるExcel
* 複数の表が1シート内にあるExcel
* セル結合を多用しているExcel
* パスワード付きExcel
* 壊れているExcelファイル
* 特殊な文字コードのCSV

---

## 設定

動作設定は `config.json` で変更できます．

### サンプル設定

```json
{
  "required_columns": [
    "company_name",
    "email",
    "url"
  ],
  "empty_check_columns": [
    "company_name",
    "email",
    "url"
  ],
  "duplicate_check_columns": [
    "company_name",
    "url"
  ],
  "email_columns": [
    "email"
  ],
  "url_columns": [
    "url",
    "website"
  ],
  "numeric_columns": [
    "price",
    "amount",
    "count",
    "quantity"
  ],
  "check_empty_cells": true,
  "check_duplicate_rows": true,
  "check_email_format": true,
  "check_url_format": true,
  "check_numeric_format": true
}
```

### 設定項目

| 項目                      | 内容                   |
| ----------------------- | -------------------- |
| required_columns        | ファイル内に必ず存在してほしい列     |
| empty_check_columns     | 空欄チェックの対象にする列        |
| duplicate_check_columns | 重複判定に使用する列           |
| email_columns           | メールアドレス形式をチェックする列    |
| url_columns             | URL形式をチェックする列        |
| numeric_columns         | 数値として扱う列             |
| check_empty_cells       | 空欄セルをチェックするかどうか      |
| check_duplicate_rows    | 重複行をチェックするかどうか       |
| check_email_format      | メールアドレス形式をチェックするかどうか |
| check_url_format        | URL形式をチェックするかどうか     |
| check_numeric_format    | 数値列をチェックするかどうか       |

---

## 設定例

### 企業リストをチェックする場合

```json
{
  "required_columns": [
    "company_name",
    "email",
    "url"
  ],
  "empty_check_columns": [
    "company_name",
    "email",
    "url"
  ],
  "duplicate_check_columns": [
    "company_name",
    "url"
  ],
  "email_columns": [
    "email"
  ],
  "url_columns": [
    "url"
  ],
  "numeric_columns": [],
  "check_empty_cells": true,
  "check_duplicate_rows": true,
  "check_email_format": true,
  "check_url_format": true,
  "check_numeric_format": true
}
```

### 商品リストをチェックする場合

```json
{
  "required_columns": [
    "product_name",
    "price",
    "url"
  ],
  "empty_check_columns": [
    "product_name",
    "price",
    "url"
  ],
  "duplicate_check_columns": [
    "product_name",
    "url"
  ],
  "email_columns": [],
  "url_columns": [
    "url"
  ],
  "numeric_columns": [
    "price",
    "stock",
    "quantity"
  ],
  "check_empty_cells": true,
  "check_duplicate_rows": true,
  "check_email_format": true,
  "check_url_format": true,
  "check_numeric_format": true
}
```

### 空欄チェックを全列に対して行う場合

`empty_check_columns` を空配列にします．

```json
{
  "empty_check_columns": []
}
```

この場合，ファイル内のすべての列が空欄チェック対象になります．

### 行全体で重複判定する場合

`duplicate_check_columns` を空配列にします．

```json
{
  "duplicate_check_columns": []
}
```

この場合，行全体が完全一致している行を重複として判定します．

---

## 出力ファイル

実行後，`output` フォルダに以下の形式でExcelファイルが出力されます．

```text
validation_report_YYYYMMDD_HHMMSS.xlsx
```

例：

```text
validation_report_20260612_150932.xlsx
```

ファイル名に実行日時が入るため，過去のレポートを上書きしにくくなっています．

---

## Excelシート

出力されるExcelファイルには，以下のシートが含まれます．

```text
summary
file_summary
missing_columns
empty_cells
duplicate_rows
invalid_values
errors
config
```

---

## summary

検査結果の全体概要を出力します．

| item            | 内容                  |
| --------------- | ------------------- |
| checked_files   | 検査対象ファイル数           |
| passed_files    | 問題がなかったファイル数        |
| failed_files    | 検査できたが不備が見つかったファイル数 |
| error_files     | 読み込みや検査処理に失敗したファイル数 |
| checked_rows    | 検査した合計行数            |
| missing_columns | 不足していた必須列の件数        |
| empty_cells     | 空欄セルの件数             |
| duplicate_rows  | 重複行の件数              |
| invalid_values  | 形式不備の件数             |
| errors          | 読み込み・検査エラー件数        |
| total_issues    | 検出された不備の合計件数        |

最初にこのシートを見ることで，全体として問題があるかどうかを確認できます．

---

## file_summary

ファイルごとの検査結果を出力します．

| Column                | Description        |
| --------------------- | ------------------ |
| file_name             | 検査対象ファイル名          |
| status                | ファイル単位の判定          |
| rows                  | 読み込んだ行数            |
| columns               | 読み込んだ列数            |
| missing_columns_count | 不足していた必須列の件数       |
| empty_cells_count     | 空欄セルの件数            |
| duplicate_rows_count  | 重複行の件数             |
| invalid_values_count  | 形式不備の件数            |
| total_issues          | ファイル内で見つかった不備の合計件数 |

### status の意味

| status | 意味                     |
| ------ | ---------------------- |
| OK     | ファイルを読み込めて，不備が見つからなかった |
| NG     | ファイルを読み込めたが，不備が見つかった   |
| ERROR  | ファイルの読み込み，または検査処理に失敗した |

---

## missing_columns

必須列が不足しているファイルを出力します．

| Column         | Description |
| -------------- | ----------- |
| file_name      | 対象ファイル名     |
| missing_column | 不足している列名    |

例：

| file_name     | missing_column |
| ------------- | -------------- |
| companies.csv | email          |
| companies.csv | url            |

`config.json` の `required_columns` に指定した列がファイル内に存在しない場合，このシートに出力されます．

---

## empty_cells

空欄セルを出力します．

| Column      | Description |
| ----------- | ----------- |
| file_name   | 対象ファイル名     |
| row_number  | ファイル上の行番号   |
| column_name | 空欄だった列名     |

例：

| file_name     | row_number | column_name  |
| ------------- | ---------- | ------------ |
| companies.csv | 6          | company_name |
| companies.csv | 7          | email        |

`row_number` は，CSV・Excel上の実際の行番号に合わせています．
1行目はヘッダーとして扱うため，データ1行目は `row_number = 2` になります．

---

## duplicate_rows

重複行を出力します．

| Column           | Description      |
| ---------------- | ---------------- |
| file_name        | 対象ファイル名          |
| row_number       | 重複している行番号        |
| duplicate_of_row | 最初に出現した同一データの行番号 |

例：

| file_name     | row_number | duplicate_of_row |
| ------------- | ---------- | ---------------- |
| companies.csv | 11         | 2                |

この例では，11行目のデータが2行目のデータと重複していることを意味します．

重複判定に使用する列は，`config.json` の `duplicate_check_columns` で指定できます．

---

## invalid_values

メールアドレス，URL，数値列などの形式不備を出力します．

| Column      | Description |
| ----------- | ----------- |
| file_name   | 対象ファイル名     |
| row_number  | ファイル上の行番号   |
| column_name | 不備があった列名    |
| value       | 実際に入っていた値   |
| issue       | 不備の種類       |

### issue の種類

| issue                 | 内容            |
| --------------------- | ------------- |
| invalid email format  | メールアドレス形式ではない |
| invalid url format    | URL形式ではない     |
| invalid numeric value | 数値として扱えない     |

例：

| file_name     | row_number | column_name | value                                     | issue                 |
| ------------- | ---------- | ----------- | ----------------------------------------- | --------------------- |
| companies.csv | 4          | email       | wrong-email                               | invalid email format  |
| companies.csv | 5          | url         | [www.example.com](http://www.example.com) | invalid url format    |
| companies.csv | 9          | price       | abc                                       | invalid numeric value |

空欄セルは `empty_cells` シートで検出するため，`invalid_values` では空欄を形式不備として扱いません．

---

## errors

ファイルの読み込み，または検査処理に失敗した場合に出力します．

| Column        | Description |
| ------------- | ----------- |
| file_name     | 対象ファイル名     |
| error_message | エラー内容       |

通常，問題なくファイルを読み込めた場合，このシートは空です．

`errors` シートが空の場合は，読み込みエラーがなかったことを意味します．

以下のような場合に，このシートへ記録される可能性があります．

* CSVの文字コードが合わず読み込めない
* Excelファイルが壊れている
* パスワード付きExcelで読み込めない
* 拡張子は `.xlsx` だが中身がExcel形式ではない
* 想定外の形式で pandas が読み込みに失敗した

---

## config

検査時に使用した設定を出力します．

| Column | Description |
| ------ | ----------- |
| key    | 設定項目        |
| value  | 設定値         |

このシートを見ることで，どの列を必須列として扱ったか，どの列をメールアドレスやURLとしてチェックしたかを確認できます．

レポートを後から見返す場合や，他の人に共有する場合に役立ちます．

---

## サンプルファイル

初回実行時，以下のファイルが存在しない場合は自動作成されます．

```text
config.json
input/sample_companies.csv
```

`sample_companies.csv` には，動作確認用に以下のような不備が含まれています．

* メールアドレス形式不正
* URL形式不正
* 空欄セル
* 数値列の不正値
* 重複行

そのため，初回実行だけで各シートの出力内容を確認できます．

---

## 注意事項

* CSV・Excelファイルは，1行目に列名がある表形式を想定しています．
* Excelファイルの場合，表はA1セルから始まる形式を推奨します．
* Excelの複数シートには対応していません．現在は最初のシートを読み込みます．
* 空欄セルは `empty_cells` シートに出力されます．
* 空欄セルは `invalid_values` では形式不備として扱いません．
* `errors` シートが空の場合は，読み込みエラーがなかったことを意味します．
* `output` フォルダ内のレポートは実行結果なので，Git管理対象から外しても問題ありません．
* `config.json` が既に存在する場合，自動作成処理では上書きされません．

---

## 活用方法

* 納品前のデータチェック
* 企業リストの品質確認
* 顧客リストの重複確認
* メールアドレス一覧の形式確認
* URL一覧の形式確認
* 商品リストの価格列チェック
* CSV・Excelの検査レポート作成
* 手作業チェックの自動化
* データ整理案件の補助ツール

---

## Requirements

```text
pandas
xlsxwriter
openpyxl
```

---

# English

## Overview

CSV Excel Data Validator is a Python tool that checks CSV and Excel files for common data quality issues and exports the results to an Excel report.

It can detect missing required columns, empty cells, duplicate rows, invalid email formats, invalid URL formats, and invalid numeric values.

The output report includes summary sheets, detailed issue sheets, read error information, and the configuration used for validation.

---

## Features

* CSV file validation
* Excel file validation
* Required column check
* Empty cell check
* Duplicate row check
* Email format check
* URL format check
* Numeric value check
* File-level OK / NG / ERROR status
* Excel report export
* Validation configuration export
* Auto-adjusted Excel column width
* Excel header formatting and filters

---

## Input Format

This tool assumes that each CSV or Excel file is formatted as a simple table.

* The first row must contain column names.
* Data rows must start from the second row.
* For Excel files, the table should start from cell A1.
* Files with title rows, blank rows before the header, or tables starting from another cell may not be read correctly.

---

## Configuration

Validation settings can be changed in `config.json`.

Example:

```json
{
  "required_columns": [
    "company_name",
    "email",
    "url"
  ],
  "empty_check_columns": [
    "company_name",
    "email",
    "url"
  ],
  "duplicate_check_columns": [
    "company_name",
    "url"
  ],
  "email_columns": [
    "email"
  ],
  "url_columns": [
    "url",
    "website"
  ],
  "numeric_columns": [
    "price",
    "amount",
    "count",
    "quantity"
  ],
  "check_empty_cells": true,
  "check_duplicate_rows": true,
  "check_email_format": true,
  "check_url_format": true,
  "check_numeric_format": true
}
```

---

## Output Sheets

The Excel report contains the following sheets.

* summary
* file_summary
* missing_columns
* empty_cells
* duplicate_rows
* invalid_values
* errors
* config

---

## Use Cases

* Data quality check before delivery
* Company list validation
* Customer list validation
* Duplicate data detection
* Email address validation
* URL validation
* Product list validation
* CSV / Excel workflow automation

---

## Limitations

* The input file must have column names in the first row.
* Excel tables should start from cell A1.
* Multiple Excel sheets are not supported in this version.
* Password-protected or broken Excel files may not be read.
* Some CSV encodings may not be read correctly.

---
