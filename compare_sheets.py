"""
Excel Sheet Comparison Tool

Compares two sheets ('N' and 'O') within an Excel workbook and highlights differences:
  - Yellow:  Row exists only in N (Added)
  - Grey:    Row exists only in O (Removed)
  - Orange:  Cell value changed between O and N
  - Green:   Column header exists only in N (New Column)

Usage:
  python compare_sheets.py <file_path> [--output-dir DIR] [--header-row 5] [--key-col J]
"""

import argparse
import os
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter


# ---------------------------------------------------------------------------
# Style constants
# ---------------------------------------------------------------------------
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GREY_FILL = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
GREEN_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _normalise(value):
    """Return a canonical string form for comparison."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value).strip()


def _values_equal(val_n, val_o):
    """Compare two cell values with tolerance for numeric formatting."""
    sn = _normalise(val_n)
    so = _normalise(val_o)
    if sn == so:
        return True
    # Try numeric comparison (e.g. "1.0" vs "1")
    try:
        if float(sn) == float(so):
            return True
    except (ValueError, TypeError):
        pass
    return False


def _read_headers(sheet, header_row):
    """Return a list of stripped header strings (None stays None)."""
    return [
        str(cell.value).strip() if cell.value is not None else None
        for cell in sheet[header_row]
    ]


def _build_data_map(sheet, start_row, key_col_idx):
    """
    Build {key_str: {header_name: value}} and {key_str: row_index} from sheet.
    key_col_idx is 0-based.
    """
    headers = _read_headers(sheet, start_row - 1)  # header row = start_row - 1
    data = {}
    key_to_row = {}
    for r_idx, row in enumerate(sheet.iter_rows(min_row=start_row), start=start_row):
        key_cell = row[key_col_idx].value
        if key_cell is None:
            continue
        key_str = str(key_cell).strip()
        if key_str == "":
            continue
        row_dict = {}
        for c_idx, cell in enumerate(row):
            col_name = headers[c_idx] if c_idx < len(headers) else f"Col_{c_idx + 1}"
            if col_name is not None:
                row_dict[col_name] = cell.value
        data[key_str] = row_dict
        key_to_row[key_str] = r_idx
    return data, key_to_row


def _header_col_map(headers):
    """Return {header_name: 1-based column index}."""
    mapping = {}
    for i, h in enumerate(headers):
        if h is not None:
            mapping[h] = i + 1  # 1-based for openpyxl
    return mapping


# ---------------------------------------------------------------------------
# Core comparison
# ---------------------------------------------------------------------------
def compare_sheets(file_path, output_dir, header_row=5, key_col="J"):
    if not os.path.isfile(file_path):
        print(f"Error: File not found: {file_path}")
        return None

    os.makedirs(output_dir, exist_ok=True)

    key_col_idx = column_index_from_string(key_col.upper()) - 1  # 0-based

    print(f"Loading {file_path} ...")
    wb_vals = openpyxl.load_workbook(file_path, data_only=True)
    wb_styles = openpyxl.load_workbook(file_path)

    for name in ("N", "O"):
        if name not in wb_vals.sheetnames:
            print(f"Error: Sheet '{name}' not found. Available: {wb_vals.sheetnames}")
            return None

    sheet_n_vals = wb_vals["N"]
    sheet_o_vals = wb_vals["O"]
    sheet_n_styles = wb_styles["N"]
    sheet_o_styles = wb_styles["O"]

    start_row = header_row + 1

    # ---- Headers ----
    n_headers = _read_headers(sheet_n_vals, header_row)
    o_headers = _read_headers(sheet_o_vals, header_row)
    o_headers_set = {h for h in o_headers if h is not None}
    n_headers_set = {h for h in n_headers if h is not None}

    n_col_map = _header_col_map(n_headers)
    o_col_map = _header_col_map(o_headers)

    # Columns present in both sheets (for value comparison)
    common_headers = [h for h in n_headers if h is not None and h in o_headers_set]

    # ---- 1. Highlight new columns in N (Green) ----
    new_columns = []
    for c_idx, h in enumerate(n_headers):
        if h is not None and h not in o_headers_set:
            sheet_n_styles.cell(row=header_row, column=c_idx + 1).fill = GREEN_FILL
            new_columns.append(h)
    if new_columns:
        print(f"  New columns in N: {', '.join(new_columns)}")

    # ---- 2. Build data maps keyed by header name ----
    print("Mapping data ...")
    n_data, n_key_to_row = _build_data_map(sheet_n_vals, start_row, key_col_idx)
    o_data, o_key_to_row = _build_data_map(sheet_o_vals, start_row, key_col_idx)

    keys_n = set(n_data)
    keys_o = set(o_data)

    diff_list = []

    # ---- 3. Added rows (in N only) -> Yellow ----
    added_keys = sorted(keys_n - keys_o)
    for key in added_keys:
        r_idx = n_key_to_row[key]
        for c in range(1, len(n_headers) + 1):
            sheet_n_styles.cell(row=r_idx, column=c).fill = YELLOW_FILL
        diff_list.append(["Added", key, "Whole Row", "(Missing in O)", "Present in N"])

    # ---- 4. Removed rows (in O only) -> Grey ----
    removed_keys = sorted(keys_o - keys_n)
    for key in removed_keys:
        r_idx = o_key_to_row[key]
        for c in range(1, len(o_headers) + 1):
            sheet_o_styles.cell(row=r_idx, column=c).fill = GREY_FILL
        diff_list.append(["Removed", key, "Whole Row", "Present in O", "(Missing in N)"])

    # ---- 5. Changed values (common rows) -> Orange ----
    changed_count = 0
    for key in sorted(keys_n & keys_o):
        vals_n = n_data[key]
        vals_o = o_data[key]
        r_idx_n = n_key_to_row[key]

        for col_name in common_headers:
            # Skip the key column itself
            if col_name == n_headers[key_col_idx]:
                continue
            v_n = vals_n.get(col_name)
            v_o = vals_o.get(col_name)

            if not _values_equal(v_n, v_o):
                col_idx_n = n_col_map[col_name]
                sheet_n_styles.cell(row=r_idx_n, column=col_idx_n).fill = ORANGE_FILL
                diff_list.append([
                    "Changed", key, col_name,
                    _normalise(v_o), _normalise(v_n),
                ])
                changed_count += 1

    # ---- 6. Summary sheet ----
    print("Creating Summary sheet ...")
    summary_name = "Summary"
    if summary_name in wb_styles.sheetnames:
        del wb_styles[summary_name]
    ws_sum = wb_styles.create_sheet(summary_name, 0)

    sum_headers = ["Type", "Key (Col {})".format(key_col), "Field", "Old Value (O)", "New Value (N)"]
    for i, h in enumerate(sum_headers, 1):
        c = ws_sum.cell(row=1, column=i, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")

    for i, entry in enumerate(diff_list, 2):
        for j, val in enumerate(entry, 1):
            ws_sum.cell(row=i, column=j, value=val)

    # Totals
    gap = len(diff_list) + 3
    stats = [
        ("Added rows:", len(added_keys)),
        ("Removed rows:", len(removed_keys)),
        ("Changed cells:", changed_count),
        ("New columns:", len(new_columns)),
        ("Total differences:", len(diff_list)),
    ]
    for offset, (label, count) in enumerate(stats):
        row = gap + offset
        ws_sum.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws_sum.cell(row=row, column=2, value=count).font = Font(bold=True)

    # Auto-fit column widths (estimate)
    for col_idx in range(1, len(sum_headers) + 1):
        max_len = len(str(sum_headers[col_idx - 1]))
        for row_idx in range(2, min(len(diff_list) + 2, 200)):
            cell_val = ws_sum.cell(row=row_idx, column=col_idx).value
            if cell_val is not None:
                max_len = max(max_len, len(str(cell_val)))
        ws_sum.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

    # ---- 7. Save ----
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_filename = f"compare_result_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    wb_styles.save(output_path)
    print(
        f"\nDone!  Added={len(added_keys)}  Removed={len(removed_keys)}  "
        f"Changed={changed_count}  NewCols={len(new_columns)}"
    )
    print(f"Saved to: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Compare sheets 'N' and 'O' in an Excel workbook and highlight differences.",
    )
    parser.add_argument("file", help="Path to the Excel (.xlsx) file")
    parser.add_argument(
        "--output-dir", "-o",
        default=None,
        help="Directory for the result file (defaults to same directory as input)",
    )
    parser.add_argument(
        "--header-row",
        type=int,
        default=5,
        help="Row number containing column headers (default: 5)",
    )
    parser.add_argument(
        "--key-col",
        default="J",
        help="Column letter used as the unique row key (default: J)",
    )
    args = parser.parse_args()

    output_dir = args.output_dir or os.path.dirname(os.path.abspath(args.file))
    compare_sheets(args.file, output_dir, header_row=args.header_row, key_col=args.key_col)


if __name__ == "__main__":
    main()
