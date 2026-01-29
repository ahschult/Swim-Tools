import os
import openpyxl
import xlrd
from openpyxl.cell.cell import TYPE_STRING, TYPE_NUMERIC, TYPE_BOOL, TYPE_NULL, TYPE_FORMULA, TYPE_ERROR

def get_excel_type_xlsx(cell):
    """Get the Excel data type for an openpyxl cell."""
    if cell.value is None:
        return "Empty"
    if cell.data_type == TYPE_NUMERIC:
        return "Date" if cell.is_date else "Num"
    elif cell.data_type == TYPE_STRING:
        return "Text"
    elif cell.data_type == TYPE_BOOL:
        return "Bool"
    elif cell.data_type == TYPE_FORMULA:
        return "Form"
    elif cell.data_type == TYPE_ERROR:
        return "Err"
    return "Unk"

def get_excel_type_xls(sheet, row_idx, col_idx):
    """Get the Excel data type for an xlrd cell."""
    cell_type = sheet.cell_type(row_idx, col_idx)
    type_map = {0: "Empty", 1: "Text", 2: "Num", 3: "Date", 4: "Bool", 5: "Err", 6: "Blank"}
    return type_map.get(cell_type, "Unk")

def read_excel_file(filename: str, num_data_rows: int = 3):
    try:
        _, ext = os.path.splitext(filename)
        print(f"\n{'='*120}")
        print(f"Reading file: {filename}")
        print(f"{'='*120}\n")

        if ext.lower() == '.xlsx':
            workbook = openpyxl.load_workbook(filename, data_only=True)
            sheet = workbook.active
            header_cells = list(sheet[1])
            max_read = min(num_data_rows, sheet.max_row - 1)
            data_rows = [list(sheet[i]) for i in range(2, 2 + max_read)]

            print(f"{'Column Header':<25} | {'Excel Types (R2-4)':<20} | {'Python Types (R2-4)':<25} | Data Samples")
            print(f"{'-'*25}-+-{'-'*20}-+-{'-'*25}-+-{'-'*40}")

            for col_idx in range(len(header_cells)):
                header_name = str(header_cells[col_idx].value or "(Empty)")
                
                ex_types, py_types, vals = [], [], []
                for row in data_rows:
                    cell = row[col_idx]
                    ex_types.append(get_excel_type_xlsx(cell))
                    py_types.append(type(cell.value).__name__)
                    v = str(cell.value)[:12] + ".." if len(str(cell.value)) > 12 else str(cell.value)
                    vals.append(v or "None")

                print(f"{header_name[:25]:<25} | {'/'.join(ex_types):<20} | {'/'.join(py_types):<25} | {' | '.join(vals)}")

        elif ext.lower() == '.xls':
            workbook = xlrd.open_workbook(filename)
            sheet = workbook.sheet_by_index(0)
            headers = sheet.row_values(0)
            max_read = min(num_data_rows, sheet.nrows - 1)
            
            print(f"{'Column Header':<25} | {'Excel Types (R2-4)':<20} | {'Python Types (R2-4)':<25} | Data Samples")
            print(f"{'-'*25}-+-{'-'*20}-+-{'-'*25}-+-{'-'*40}")

            for col_idx in range(len(headers)):
                header_name = str(headers[col_idx] or "(Empty)")
                ex_types, py_types, vals = [], [], []
                
                for r_idx in range(1, 1 + max_read):
                    val = sheet.cell_value(r_idx, col_idx)
                    ex_types.append(get_excel_type_xls(sheet, r_idx, col_idx))
                    py_types.append(type(val).__name__)
                    v = str(val)[:12] + ".." if len(str(val)) > 12 else str(val)
                    vals.append(v if val != "" else "Empty")

                print(f"{header_name[:25]:<25} | {'/'.join(ex_types):<20} | {'/'.join(py_types):<25} | {' | '.join(vals)}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    file_path = "filename_here.xls/xlsx"  # Replace with your file path
    read_excel_file(file_path, num_data_rows=3) 