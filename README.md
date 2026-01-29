# Excel Data Inspector

A Python utility for quickly inspecting Excel file structure, data types, and sample data. This tool helps you understand your spreadsheet's schema before performing data analysis.

## Features

- Supports both `.xls` (Excel 97-2003) and `.xlsx` (Excel 2007+) formats
- Displays column headers from the first row
- Shows Excel data types (Number, Text, Date, Boolean, etc.)
- Shows Python data types (int, float, str, etc.)
- Previews first 3-4 rows of actual data
- Compact, readable output format

## Installation

### Prerequisites

Python 3.6 or higher is required.

### Required Libraries

Install the required dependencies using pip:

```bash
pip install openpyxl xlrd
```

Or if you have a `requirements.txt`:

```bash
pip install -r requirements.txt
```

**requirements.txt:**
```
openpyxl>=3.0.0
xlrd>=2.0.0
```

## Usage

### Basic Usage

1. Update the file path in the script:
```python
file_path = "path/to/your/file.xls"  # or .xlsx
```

2. Run the script:
```bash
python excel_data_inspector.py
```


## Use Cases

### 1. Data Analysis Planning
Quickly understand the structure of an unfamiliar spreadsheet before writing analysis code.

```python
read_excel_file("sales_data.xlsx", num_data_rows=5)
```

### 2. Data Type Validation
Verify that columns contain the expected data types before processing.

### 3. Schema Documentation
Generate a quick overview of your data schema for documentation purposes.

### 4. Data Quality Checks
Identify columns with inconsistent data types or unexpected values.

## Example

```python
import os
import openpyxl
import xlrd
from openpyxl.cell.cell import TYPE_STRING, TYPE_NUMERIC, TYPE_BOOL, TYPE_NULL, TYPE_FORMULA, TYPE_ERROR

# ... (include all the functions from the script)

if __name__ == "__main__":
    # Inspect a sales report
    file_path = "data/Q1_2025_Sales.xlsx"
    read_excel_file(file_path, num_data_rows=3)
```

**Sample Output:**
```
====================================================================================================================
Reading file: data/Q1_2025_Sales.xlsx
====================================================================================================================

Column Header             | Excel Types (R2-4)   | Python Types (R2-4)       | Data Samples
-------------------------+---------------------+--------------------------+------------------------------------------
(Empty)                   | Text/Text/Text       | str/str/str               | SWIMTIME | 25.66 | 26.54
(Empty)                   | Text/Num/Num         | str/float/float           | SWIMTIME_N | 25.66 | 26.54
AQUA 2025                 | Text/Num/Num         | str/float/float           | PTS_FINA | 466.0 | 421.0
```

## Troubleshooting

### File Not Found
```
Error: [Errno 2] No such file or directory: 'file.xlsx'
```
**Solution:** Check that the file path is correct and the file exists.

### Module Not Found
```
ModuleNotFoundError: No module named 'openpyxl'
```
**Solution:** Install the required library:
```bash
pip install openpyxl
```

### Unicode/Encoding Errors
If you encounter encoding issues with .xls files, the script handles them gracefully and displays what it can.

## Limitations

- Assumes the first row contains headers
- Truncates long values (>12 characters) in the sample data display
- For .xls files, dates are shown as numeric values (Excel's serial date format)
- Does not support password-protected files


## License

This script is provided as-is for educational and analytical purposes.


