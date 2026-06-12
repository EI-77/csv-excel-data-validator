import json
import os
import re
from datetime import datetime

import pandas as pd


INPUT_DIR = "input"
OUTPUT_DIR = "output"
CONFIG_FILE = "config.json"

SUPPORTED_EXTENSIONS = [".csv", ".xlsx"]


def create_sample_files():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(CONFIG_FILE):
        sample_config = {
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
            "check_empty_cells": True,
            "check_duplicate_rows": True,
            "check_email_format": True,
            "check_url_format": True,
            "check_numeric_format": True
        }

        with open(CONFIG_FILE, "w", encoding="utf-8") as file:
            json.dump(sample_config, file, ensure_ascii=False, indent=2)

        print(f"Sample config created: {CONFIG_FILE}")

    sample_csv_path = os.path.join(INPUT_DIR, "sample_companies.csv")

    if not os.path.exists(sample_csv_path):
        sample_data = [
            {
                "company_name": "ABC株式会社",
                "email": "info@abc.co.jp",
                "url": "https://www.abc.co.jp",
                "website": "https://abc.co.jp",
                "price": "10000",
                "amount": "50000",
                "count": "10",
                "quantity": "3"
            },
            {
                "company_name": "XYZ株式会社",
                "email": "contact@xyz.jp",
                "url": "https://www.xyz.jp",
                "website": "https://xyz.jp",
                "price": "25000",
                "amount": "120000",
                "count": "5",
                "quantity": "8"
            },
            {
                "company_name": "サンプル商事",
                "email": "wrong-email",
                "url": "https://www.sample.co.jp",
                "website": "https://sample.co.jp",
                "price": "15000",
                "amount": "70000",
                "count": "7",
                "quantity": "2"
            },
            {
                "company_name": "テスト工業",
                "email": "support@test.co.jp",
                "url": "www.test.co.jp",
                "website": "https://test.co.jp",
                "price": "30000",
                "amount": "90000",
                "count": "12",
                "quantity": "4"
            },
            {
                "company_name": "",
                "email": "info@blank.co.jp",
                "url": "https://www.blank.co.jp",
                "website": "https://blank.co.jp",
                "price": "18000",
                "amount": "60000",
                "count": "6",
                "quantity": "1"
            },
            {
                "company_name": "メール欠損株式会社",
                "email": "",
                "url": "https://www.noemail.co.jp",
                "website": "https://noemail.co.jp",
                "price": "22000",
                "amount": "80000",
                "count": "9",
                "quantity": "5"
            },
            {
                "company_name": "URL欠損株式会社",
                "email": "info@nourl.co.jp",
                "url": "",
                "website": "https://nourl.co.jp",
                "price": "12000",
                "amount": "40000",
                "count": "4",
                "quantity": "2"
            },
            {
                "company_name": "数値不正株式会社",
                "email": "info@invalid-number.co.jp",
                "url": "https://www.invalid-number.co.jp",
                "website": "https://invalid-number.co.jp",
                "price": "abc",
                "amount": "50000",
                "count": "10",
                "quantity": "3"
            },
            {
                "company_name": "数量不正株式会社",
                "email": "info@invalid-quantity.co.jp",
                "url": "https://www.invalid-quantity.co.jp",
                "website": "https://invalid-quantity.co.jp",
                "price": "14000",
                "amount": "unknown",
                "count": "ten",
                "quantity": "three"
            },
            {
                "company_name": "ABC株式会社",
                "email": "info@abc.co.jp",
                "url": "https://www.abc.co.jp",
                "website": "https://abc.co.jp",
                "price": "10000",
                "amount": "50000",
                "count": "10",
                "quantity": "3"
            }
        ]

        df = pd.DataFrame(sample_data)
        df.to_csv(sample_csv_path, index=False, encoding="utf-8-sig")

        print(f"Sample CSV created: {sample_csv_path}")


def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as file:
        return json.load(file)


def find_input_files():
    input_files = []

    for file_name in os.listdir(INPUT_DIR):
        file_path = os.path.join(INPUT_DIR, file_name)
        extension = os.path.splitext(file_name)[1].lower()

        if os.path.isfile(file_path) and extension in SUPPORTED_EXTENSIONS:
            input_files.append(file_path)

    return input_files


def read_table_file(file_path):
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".csv":
        return pd.read_csv(file_path)

    if extension == ".xlsx":
        return pd.read_excel(file_path)

    raise ValueError(f"Unsupported file type: {file_path}")


def is_empty(value):
    if pd.isna(value):
        return True

    if str(value).strip() == "":
        return True

    return False


def is_valid_email(value):
    if is_empty(value):
        return True

    email_pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

    return re.match(email_pattern, str(value).strip()) is not None


def is_valid_url(value):
    if is_empty(value):
        return True

    url = str(value).strip()

    return url.startswith("http://") or url.startswith("https://")


def is_valid_number(value):
    if is_empty(value):
        return True

    converted_value = pd.to_numeric(value, errors="coerce")

    return not pd.isna(converted_value)


def check_missing_columns(file_name, df, config):
    results = []

    required_columns = config.get("required_columns", [])

    for column in required_columns:
        if column not in df.columns:
            results.append({
                "file_name": file_name,
                "missing_column": column
            })

    return results


def check_empty_cells(file_name, df, config):
    results = []

    empty_check_columns = config.get("empty_check_columns", [])

    if empty_check_columns:
        target_columns = [
            column for column in empty_check_columns
            if column in df.columns
        ]
    else:
        target_columns = list(df.columns)

    for row_index, row in df.iterrows():
        for column in target_columns:
            if is_empty(row[column]):
                results.append({
                    "file_name": file_name,
                    "row_number": row_index + 2,
                    "column_name": column
                })

    return results


def check_duplicate_rows(file_name, df, config):
    results = []
    first_seen_rows = {}

    duplicate_check_columns = config.get("duplicate_check_columns", [])

    if duplicate_check_columns:
        target_columns = [
            column for column in duplicate_check_columns
            if column in df.columns
        ]
    else:
        target_columns = list(df.columns)

    if not target_columns:
        return results

    for row_index, row in df.iterrows():
        row_key = tuple(
            "" if is_empty(row[column]) else str(row[column]).strip()
            for column in target_columns
        )

        if row_key in first_seen_rows:
            results.append({
                "file_name": file_name,
                "row_number": row_index + 2,
                "duplicate_of_row": first_seen_rows[row_key] + 2
            })
        else:
            first_seen_rows[row_key] = row_index

    return results


def check_email_values(file_name, df, config):
    results = []

    email_columns = config.get("email_columns", [])

    for column in email_columns:
        if column not in df.columns:
            continue

        for row_index, value in df[column].items():
            if not is_valid_email(value):
                results.append({
                    "file_name": file_name,
                    "row_number": row_index + 2,
                    "column_name": column,
                    "value": value,
                    "issue": "invalid email format"
                })

    return results


def check_url_values(file_name, df, config):
    results = []

    url_columns = config.get("url_columns", [])

    for column in url_columns:
        if column not in df.columns:
            continue

        for row_index, value in df[column].items():
            if not is_valid_url(value):
                results.append({
                    "file_name": file_name,
                    "row_number": row_index + 2,
                    "column_name": column,
                    "value": value,
                    "issue": "invalid url format"
                })

    return results


def check_numeric_values(file_name, df, config):
    results = []

    numeric_columns = config.get("numeric_columns", [])

    for column in numeric_columns:
        if column not in df.columns:
            continue

        for row_index, value in df[column].items():
            if not is_valid_number(value):
                results.append({
                    "file_name": file_name,
                    "row_number": row_index + 2,
                    "column_name": column,
                    "value": value,
                    "issue": "invalid numeric value"
                })

    return results


def validate_file(file_path, config):
    file_name = os.path.basename(file_path)

    print(f"Checking file: {file_name}")

    df = read_table_file(file_path)

    missing_columns = check_missing_columns(file_name, df, config)

    empty_cells = []
    duplicate_rows = []
    invalid_values = []

    if config.get("check_empty_cells", True):
        empty_cells = check_empty_cells(file_name, df, config)

    if config.get("check_duplicate_rows", True):
        duplicate_rows = check_duplicate_rows(file_name, df, config)

    if config.get("check_email_format", True):
        invalid_values.extend(check_email_values(file_name, df, config))

    if config.get("check_url_format", True):
        invalid_values.extend(check_url_values(file_name, df, config))

    if config.get("check_numeric_format", True):
        invalid_values.extend(check_numeric_values(file_name, df, config))

    total_issues = (
        len(missing_columns)
        + len(empty_cells)
        + len(duplicate_rows)
        + len(invalid_values)
    )

    status = "OK" if total_issues == 0 else "NG"

    file_summary = {
        "file_name": file_name,
        "status": status,
        "rows": len(df),
        "columns": len(df.columns),
        "missing_columns_count": len(missing_columns),
        "empty_cells_count": len(empty_cells),
        "duplicate_rows_count": len(duplicate_rows),
        "invalid_values_count": len(invalid_values),
        "total_issues": total_issues
    }

    return {
        "file_summary": file_summary,
        "missing_columns": missing_columns,
        "empty_cells": empty_cells,
        "duplicate_rows": duplicate_rows,
        "invalid_values": invalid_values
    }


def make_summary(file_summaries, missing_columns, empty_cells, duplicate_rows, invalid_values, errors):
    checked_files = len(file_summaries)
    passed_files = sum(
        1 for summary in file_summaries
        if summary.get("status") == "OK"
    )
    failed_files = sum(
        1 for summary in file_summaries
        if summary.get("status") == "NG"
    )
    error_files = sum(
        1 for summary in file_summaries
        if summary.get("status") == "ERROR"
    )
    checked_rows = sum(summary["rows"] for summary in file_summaries)

    summary_df = pd.DataFrame([
        {
            "item": "checked_files",
            "value": checked_files
        },
        {
            "item": "passed_files",
            "value": passed_files
        },
        {
            "item": "failed_files",
            "value": failed_files
        },
        {
            "item": "error_files",
            "value": error_files
        },
        {
            "item": "checked_rows",
            "value": checked_rows
        },
        {
            "item": "missing_columns",
            "value": len(missing_columns)
        },
        {
            "item": "empty_cells",
            "value": len(empty_cells)
        },
        {
            "item": "duplicate_rows",
            "value": len(duplicate_rows)
        },
        {
            "item": "invalid_values",
            "value": len(invalid_values)
        },
        {
            "item": "errors",
            "value": len(errors)
        },
        {
            "item": "total_issues",
            "value": len(missing_columns) + len(empty_cells) + len(duplicate_rows) + len(invalid_values)
        }
    ])

    return summary_df


def make_config_df(config):
    rows = []

    for key, value in config.items():
        if isinstance(value, list):
            display_value = ", ".join(str(item) for item in value)
        else:
            display_value = value

        rows.append({
            "key": key,
            "value": display_value
        })

    return pd.DataFrame(rows, columns=["key", "value"])


def make_output_path():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"validation_report_{timestamp}.xlsx"

    return os.path.join(OUTPUT_DIR, file_name)


def adjust_column_width(worksheet, df):
    if len(df.columns) == 0:
        return

    for index, column in enumerate(df.columns):
        if len(df) == 0:
            max_length = len(column)
        else:
            max_length = max(
                df[column].map(lambda value: len(str(value))).max(),
                len(column)
            )

        width = min(max_length + 2, 60)
        worksheet.set_column(index, index, width)


def format_worksheet(writer, worksheet, df):
    if len(df.columns) == 0:
        return

    header_format = writer.book.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1
    })

    for col_num, column in enumerate(df.columns):
        worksheet.write(0, col_num, column, header_format)

    last_row = max(len(df), 1)
    worksheet.autofilter(0, 0, last_row, len(df.columns) - 1)
    worksheet.freeze_panes(1, 0)


def save_report(file_summaries, missing_columns, empty_cells, duplicate_rows, invalid_values, errors, config):
    output_path = make_output_path()

    summary_df = make_summary(
        file_summaries,
        missing_columns,
        empty_cells,
        duplicate_rows,
        invalid_values,
        errors
    )

    file_summary_df = pd.DataFrame(
        file_summaries,
        columns=[
            "file_name",
            "status",
            "rows",
            "columns",
            "missing_columns_count",
            "empty_cells_count",
            "duplicate_rows_count",
            "invalid_values_count",
            "total_issues"
        ]
    )

    missing_columns_df = pd.DataFrame(
        missing_columns,
        columns=[
            "file_name",
            "missing_column"
        ]
    )

    empty_cells_df = pd.DataFrame(
        empty_cells,
        columns=[
            "file_name",
            "row_number",
            "column_name"
        ]
    )

    duplicate_rows_df = pd.DataFrame(
        duplicate_rows,
        columns=[
            "file_name",
            "row_number",
            "duplicate_of_row"
        ]
    )

    invalid_values_df = pd.DataFrame(
        invalid_values,
        columns=[
            "file_name",
            "row_number",
            "column_name",
            "value",
            "issue"
        ]
    )

    errors_df = pd.DataFrame(
        errors,
        columns=[
            "file_name",
            "error_message"
        ]
    )

    config_df = make_config_df(config)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        sheets = {
            "summary": summary_df,
            "file_summary": file_summary_df,
            "missing_columns": missing_columns_df,
            "empty_cells": empty_cells_df,
            "duplicate_rows": duplicate_rows_df,
            "invalid_values": invalid_values_df,
            "errors": errors_df,
            "config": config_df
        }

        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]
            adjust_column_width(worksheet, df)
            format_worksheet(writer, worksheet, df)

    print(f"Report saved: {output_path}")


def main():
    create_sample_files()

    config = load_config()
    input_files = find_input_files()

    if not input_files:
        print(f"No CSV or Excel files found in {INPUT_DIR}.")
        return

    file_summaries = []
    all_missing_columns = []
    all_empty_cells = []
    all_duplicate_rows = []
    all_invalid_values = []
    all_errors = []

    for file_path in input_files:
        try:
            result = validate_file(file_path, config)

            file_summaries.append(result["file_summary"])
            all_missing_columns.extend(result["missing_columns"])
            all_empty_cells.extend(result["empty_cells"])
            all_duplicate_rows.extend(result["duplicate_rows"])
            all_invalid_values.extend(result["invalid_values"])

        except Exception as error:
            file_name = os.path.basename(file_path)
            error_message = str(error)

            file_summaries.append({
                "file_name": file_name,
                "status": "ERROR",
                "rows": 0,
                "columns": 0,
                "missing_columns_count": 0,
                "empty_cells_count": 0,
                "duplicate_rows_count": 0,
                "invalid_values_count": 0,
                "total_issues": 0
            })

            all_errors.append({
                "file_name": file_name,
                "error_message": error_message
            })

            print(f"Failed to check file: {file_name}")
            print(f"Error: {error_message}")

    save_report(
        file_summaries,
        all_missing_columns,
        all_empty_cells,
        all_duplicate_rows,
        all_invalid_values,
        all_errors,
        config
    )

    print(f"Checked files: {len(file_summaries)}")
    print(f"Passed files: {sum(1 for summary in file_summaries if summary.get('status') == 'OK')}")
    print(f"Failed files: {sum(1 for summary in file_summaries if summary.get('status') == 'NG')}")
    print(f"Error files: {sum(1 for summary in file_summaries if summary.get('status') == 'ERROR')}")
    print(f"Missing columns: {len(all_missing_columns)}")
    print(f"Empty cells: {len(all_empty_cells)}")
    print(f"Duplicate rows: {len(all_duplicate_rows)}")
    print(f"Invalid values: {len(all_invalid_values)}")


if __name__ == "__main__":
    main()
    