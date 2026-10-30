# Excel Data Quality Profiler

A Python-based tool that analyzes Excel workbooks, identifies common data-quality issues, and generates automated Excel profiling reports.

## Features

- **Multi-sheet profiling** – Analyzes every worksheet in an Excel workbook.
- **Missing Value Detection** – Identifies empty cells and whitespace-only values.
- **Duplicate Detection** – Detects duplicate rows across worksheets.
- **Business Key Validation** – Checks duplicate combinations of configurable business-key columns.
- **Column Profiling** – Analyzes data types, unique values, missing percentages, and sample values.
- **Numeric & Date Diagnostics** – Evaluates how successfully column values can be parsed as numbers and dates.
- **ID Column Analysis** – Identifies duplicate values in columns with ID-like names.
- **Automated Excel Reports** – Generates a separate Excel workbook containing profiling results.
- **Source Data Protection** – The original workbook is never modified.

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Tkinter

## Project Structure

```text
excel-data-quality-profiler/
│
├── excel_data_profiler.py
├── README.md
├── requirements.txt
├── .gitignore
│
├── sample_data/
│
└── screenshots/
```

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/<your-username>/excel-data-quality-profiler.git
```

**2. Navigate to the project directory**

```bash
cd excel-data-quality-profiler
```

**3. Install dependencies**

```bash
pip install -r requirements.txt
```

## Usage

Run the profiler:

```bash
python excel_data_profiler.py
```

A file picker will allow you to select an Excel workbook when Tkinter is available.

Alternatively, provide the workbook path directly:

```bash
python excel_data_profiler.py "path/to/workbook.xlsx"
```

Specify a custom output directory:

```bash
python excel_data_profiler.py "path/to/workbook.xlsx" --output-dir "profiling_output"
```

Configure the number of sample values displayed per column:

```bash
python excel_data_profiler.py "path/to/workbook.xlsx" --sample-size 10
```

## Business Key Configuration

The profiler supports configurable business-key duplicate checks.

Edit the `BUSINESS_KEY_COLUMNS` dictionary inside the Python script:

```python
BUSINESS_KEY_COLUMNS = {
    "Sales": ["Order ID", "Product ID"],
    "Customers": ["Customer ID"]
}
```

Replace the example worksheet and column names with those present in your workbook.

## Generated Report

The profiler generates a separate Excel report containing six worksheets:

| Worksheet | Description |
|---|---|
| Sheet_Summary | Overall worksheet statistics |
| Column_Profile | Column-level profiling information |
| Missing_Values | Missing-value statistics |
| Duplicate_IDs | Duplicate values in ID-like columns |
| Business_Key_Duplicates | Business-key duplicate statistics |
| Quality_Issues | Identified data-quality issues |

The generated report is saved in the specified output directory.

## Supported File Formats

- `.xlsx`
- `.xls`
- `.xlsm`

## Notes

- The source workbook is not modified.
- Business-key validation requires manual configuration.
- Numeric and date parsing are diagnostic checks and do not modify original values.
- Sample datasets and screenshots can be added to their respective directories.

## Author

**Arun Kumar**

- [GitHub](https://github.com/ArunKumarI2004I)
- [LinkedIn](https://www.linkedin.com/in/arun-kumar-982a13243)

---

*Developed as a Python project focused on Excel data profiling and data-quality analysis.*
