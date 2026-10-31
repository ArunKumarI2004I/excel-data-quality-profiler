"""
Excel and CSV Data Quality Profiler
-----------------------------------
Profiles every worksheet in an Excel workbook or a single CSV file and
writes a separate Excel report.

Features
- Multi-sheet Excel workbook and single-table CSV profiling
- Blank/whitespace missing-value detection
- Full-row duplicate detection
- Configurable business-key duplicate checks
- Column type, cardinality, and sample-value profiling
- Numeric and date parsing diagnostics
- Separate Excel report; source workbook is never modified

Usage:
    python excel_data_profiler.py
    python excel_data_profiler.py "path/to/workbook.xlsx"
    python excel_data_profiler.py "path/to/data.csv"
    python excel_data_profiler.py "path/to/data.csv" --output-dir "profiling_output"

Optional configuration:
    Edit BUSINESS_KEY_COLUMNS below to define business keys by sheet.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd


# Add business-key definitions here when you know the relevant columns.
# Example:
# BUSINESS_KEY_COLUMNS = {
#     "Sales": ["Order ID", "Product ID"],
#     "Customers": ["Customer ID"],
# }
BUSINESS_KEY_COLUMNS: dict[str, list[str]] = {}

EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Profile data quality in an Excel workbook or CSV file."
    )
    parser.add_argument(
        "workbook",
        nargs="?",
        help="Path to an Excel workbook or CSV file. If omitted, a file picker or prompt is used.",
    )
    parser.add_argument(
        "--output-dir",
        default="profiling_output",
        help="Directory for the generated report (default: profiling_output).",
    )
    parser.add_argument(
        "--sample-size",
        type=int,
        default=5,
        help="Maximum number of distinct sample values to show per column (default: 5).",
    )
    return parser.parse_args()


def choose_workbook(provided_path: str | None) -> Path:
    if provided_path:
        workbook = Path(provided_path).expanduser()
    else:
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            selected = filedialog.askopenfilename(
                title="Select an Excel workbook or CSV file to profile",
                filetypes=[
                    ("Excel and CSV files", "*.xlsx *.xls *.xlsm *.csv"),
                    ("All files", "*.*"),
                ],
            )
            root.destroy()
            if not selected:
                raise SystemExit("No workbook selected. Exiting.")
            workbook = Path(selected)
        except (ImportError, ModuleNotFoundError):
            entered = input("Enter the path to the Excel workbook or CSV file: ").strip().strip('"')
            if not entered:
                raise SystemExit("No workbook path provided. Exiting.")
            workbook = Path(entered)

    if not workbook.is_file():
        raise FileNotFoundError(f"Workbook not found: {workbook.resolve()}")
    if workbook.suffix.lower() not in EXCEL_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type: {workbook.suffix}. "
            f"Expected one of: {', '.join(sorted(EXCEL_EXTENSIONS))}"
        )
    return workbook


def normalize_missing(df: pd.DataFrame) -> pd.DataFrame:
    """Treat empty strings and whitespace-only strings as missing."""
    return df.replace(r"^\s*$", pd.NA, regex=True)


def safe_samples(series: pd.Series, sample_size: int) -> str:
    values = series.dropna().astype(str).drop_duplicates().head(sample_size).tolist()
    return " | ".join(values)


def profile_column(
    sheet_name: str,
    column: Any,
    series: pd.Series,
    normalized_series: pd.Series,
    row_count: int,
    sample_size: int,
) -> dict[str, Any]:
    non_missing = normalized_series.dropna()
    missing_count = int(normalized_series.isna().sum())
    unique_count = int(non_missing.nunique(dropna=True))

    # Preserve pandas' inferred dtype while also reporting likely parseability.
    inferred_dtype = str(series.dtype)
    numeric_parse_rate = None
    date_parse_rate = None

    if len(non_missing):
        numeric_values = pd.to_numeric(non_missing, errors="coerce")
        numeric_parse_rate = round(float(numeric_values.notna().mean() * 100), 2)

        # Date parsing is diagnostic only; it does not alter the source data.
        date_values = pd.to_datetime(non_missing, errors="coerce", format="mixed")
        date_parse_rate = round(float(date_values.notna().mean() * 100), 2)

    return {
        "sheet": sheet_name,
        "column": str(column),
        "inferred_dtype": inferred_dtype,
        "row_count": row_count,
        "non_missing_count": int(len(non_missing)),
        "missing_count": missing_count,
        "missing_percent": round(missing_count / row_count * 100, 2) if row_count else 0,
        "unique_values": unique_count,
        "uniqueness_percent_of_non_missing": (
            round(unique_count / len(non_missing) * 100, 2) if len(non_missing) else 0
        ),
        "numeric_parse_percent": numeric_parse_rate,
        "date_parse_percent": date_parse_rate,
        "sample_values": safe_samples(non_missing, sample_size),
    }


def load_input_file(input_path: Path) -> dict[str, pd.DataFrame]:
    """Load an Excel workbook or CSV file into a sheet-name/DataFrame mapping."""
    if input_path.suffix.lower() == ".csv":
        # Keep the same profiling workflow for CSV: treat it as one logical sheet.
        # utf-8-sig also handles CSV files saved with a UTF-8 BOM (common in Excel).
        try:
            df = pd.read_csv(input_path, dtype=object, encoding="utf-8-sig")
        except UnicodeDecodeError:
            # Fall back for CSV files exported using a Windows/Excel legacy encoding.
            df = pd.read_csv(input_path, dtype=object, encoding="cp1252")
        return {input_path.stem: df}

    # Read values as objects first so mixed columns are not silently coerced to strings.
    return pd.read_excel(input_path, sheet_name=None, dtype=object)


def profile_workbook(
    workbook: Path, output_dir: Path, sample_size: int
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    sheets = load_input_file(workbook)

    sheet_summary: list[dict[str, Any]] = []
    column_profile: list[dict[str, Any]] = []
    missing_report: list[dict[str, Any]] = []
    duplicate_report: list[dict[str, Any]] = []
    key_duplicate_report: list[dict[str, Any]] = []
    quality_issues: list[dict[str, Any]] = []

    for sheet_name, df in sheets.items():
        print(f"\n{'=' * 60}\nSHEET: {sheet_name}\n{'=' * 60}")
        row_count, column_count = df.shape
        normalized = normalize_missing(df)
        duplicate_rows = int(normalized.duplicated().sum())

        sheet_summary.append(
            {
                "sheet": sheet_name,
                "row_count": row_count,
                "column_count": column_count,
                "duplicate_full_rows": duplicate_rows,
                "empty_rows": int(normalized.isna().all(axis=1).sum()) if column_count else 0,
            }
        )

        print(f"Rows: {row_count:,}")
        print(f"Columns: {column_count:,}")
        print(f"Duplicate full rows: {duplicate_rows:,}")

        for column in df.columns:
            normalized_series = normalized[column]
            col_profile = profile_column(
                sheet_name,
                column,
                df[column],
                normalized_series,
                row_count,
                sample_size,
            )
            column_profile.append(col_profile)

            missing_report.append(
                {
                    "sheet": sheet_name,
                    "column": str(column),
                    "missing_count": col_profile["missing_count"],
                    "missing_percent": col_profile["missing_percent"],
                }
            )

            if col_profile["missing_count"]:
                quality_issues.append(
                    {
                        "sheet": sheet_name,
                        "column": str(column),
                        "issue": "Missing values",
                        "count": col_profile["missing_count"],
                        "details": f'{col_profile["missing_percent"]}% of rows are missing',
                    }
                )

            print(
                f"  {column}: dtype={col_profile['inferred_dtype']}, "
                f"missing={col_profile['missing_count']}, "
                f"unique={col_profile['unique_values']}"
            )

        if duplicate_rows:
            quality_issues.append(
                {
                    "sheet": sheet_name,
                    "column": "(entire row)",
                    "issue": "Duplicate full rows",
                    "count": duplicate_rows,
                    "details": "Rows duplicate an earlier complete row.",
                }
            )

        # Check configured business keys, even if their names do not end in "_id".
        key_columns = BUSINESS_KEY_COLUMNS.get(sheet_name, [])
        if key_columns:
            absent = [col for col in key_columns if col not in normalized.columns]
            if absent:
                quality_issues.append(
                    {
                        "sheet": sheet_name,
                        "column": ", ".join(absent),
                        "issue": "Configured business key column not found",
                        "count": len(absent),
                        "details": "Review BUSINESS_KEY_COLUMNS configuration.",
                    }
                )
            else:
                key_frame = normalized[key_columns]
                complete_key_mask = key_frame.notna().all(axis=1)
                duplicate_mask = key_frame.duplicated(keep="first") & complete_key_mask
                duplicate_count = int(duplicate_mask.sum())

                key_duplicate_report.append(
                    {
                        "sheet": sheet_name,
                        "business_key": " + ".join(key_columns),
                        "duplicate_occurrences": duplicate_count,
                        "rows_with_incomplete_key": int((~complete_key_mask).sum()),
                    }
                )

                if duplicate_count:
                    quality_issues.append(
                        {
                            "sheet": sheet_name,
                            "column": " + ".join(key_columns),
                            "issue": "Duplicate business key",
                            "count": duplicate_count,
                            "details": "Repeated complete business-key combinations.",
                        }
                    )

        # Generic duplicate-value diagnostics for columns whose names look like IDs.
        for column in df.columns:
            normalized_name = str(column).strip().lower().replace(" ", "").replace("-", "")
            if normalized_name.endswith("id") or normalized_name.endswith("_key"):
                values = normalized[column].dropna()
                duplicate_count = int(values.duplicated(keep="first").sum())
                duplicate_report.append(
                    {
                        "sheet": sheet_name,
                        "column": str(column),
                        "duplicate_occurrences": duplicate_count,
                        "unique_non_missing_values": int(values.nunique()),
                    }
                )

        # Basic numeric/date parseability warnings for columns that look numeric/date-like.
        for item in column_profile[-column_count:] if column_count else []:
            col_name = item["column"].lower()
            if any(token in col_name for token in ("amount", "price", "revenue", "quantity", "units", "total")):
                rate = item["numeric_parse_percent"]
                if rate is not None and rate < 100:
                    quality_issues.append(
                        {
                            "sheet": sheet_name,
                            "column": item["column"],
                            "issue": "Non-numeric values in numeric-looking column",
                            "count": None,
                            "details": f"Only {rate}% of non-missing values parse as numeric.",
                        }
                    )
            if any(token in col_name for token in ("date", "time")):
                rate = item["date_parse_percent"]
                if rate is not None and rate < 100:
                    quality_issues.append(
                        {
                            "sheet": sheet_name,
                            "column": item["column"],
                            "issue": "Unparseable values in date-looking column",
                            "count": None,
                            "details": f"Only {rate}% of non-missing values parse as dates.",
                        }
                    )

    report_path = output_dir / f"{workbook.stem}_data_quality_report.xlsx"
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        report_sheets = {
            "Sheet_Summary": sheet_summary,
            "Column_Profile": column_profile,
            "Missing_Values": missing_report,
            "Duplicate_IDs": duplicate_report,
            "Business_Key_Duplicates": key_duplicate_report,
            "Quality_Issues": quality_issues,
        }

        for sheet_name, records in report_sheets.items():
            pd.DataFrame(records).to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            if ws.max_row and ws.max_column:
                ws.auto_filter.ref = ws.dimensions
            for cells in ws.columns:
                max_length = max(
                    (len(str(cell.value)) if cell.value is not None else 0)
                    for cell in cells
                )
                ws.column_dimensions[cells[0].column_letter].width = min(
                    max(max_length + 2, 12), 45
                )
            for cell in ws[1]:
                cell.font = __import__("openpyxl").styles.Font(bold=True)

    print("\nProfiling complete!")
    print(f"Report saved to: {report_path.resolve()}")
    print("The source file was not modified.")
    return report_path


def main() -> None:
    args = parse_args()
    if args.sample_size < 1:
        raise ValueError("--sample-size must be at least 1.")
    input_file = choose_workbook(args.workbook)
    profile_workbook(input_file, Path(args.output_dir), args.sample_size)


if __name__ == "__main__":
    main()
