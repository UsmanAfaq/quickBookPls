"""Single-file feature: add a "(USD)" column next to every data column."""
import io
from copy import copy

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from .sheet_reader import (column_header_info, find_companies, find_row_label,
                           find_total_column)

USD_FMT = '"$"#,##0.00_);\\("$"#,##0.00\\)'


def copy_style(src_cell, dst_cell):
    """Carry a cell's look over (works across workbooks)."""
    dst_cell.font = copy(src_cell.font)
    dst_cell.border = copy(src_cell.border)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format


def add_usd_column(file_bytes: bytes, rate: float):
    """Returns (output_bytes, summary_rows)."""
    wb = load_workbook(io.BytesIO(file_bytes))
    wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True)

    ws = ws_values = None
    total_col = header_row = None
    for sheet, sheet_values in zip(wb.worksheets, wb_values.worksheets):
        col, row = find_total_column(sheet)
        if col:
            ws, ws_values, total_col, header_row = sheet, sheet_values, col, row
            break
    if ws is None:
        raise ValueError('Could not find a "TOTAL" header on any sheet.')

    companies = find_companies(ws, total_col, header_row)
    data_cols = sorted([c["col"] for c in companies] + [total_col])

    font_bold = Font(name="Arial", size=8, bold=True, color="FF323232")
    num_fmt = USD_FMT

    offset = 0
    usd_cols = []
    for orig_col in data_cols:
        actual_col = orig_col + offset
        header_text, name_row = column_header_info(ws, actual_col, header_row)
        ws.insert_cols(actual_col + 1)
        usd_cols.append((orig_col, actual_col, actual_col + 1, header_text, name_row))
        offset += 1

    rate_ref = f"{rate:.10g}"

    summary = []
    for orig_col, data_col, usd_col, header_text, name_row in usd_cols:
        is_grand_total = header_text.strip().lower() == "total"
        data_letter = get_column_letter(data_col)
        usd_header = f"{header_text} (USD)" if header_text else "USD"

        for hr in range(1, max(header_row, 2) + 1):
            src_head = ws.cell(row=hr, column=data_col)
            if hr == name_row:
                hc = ws.cell(row=hr, column=usd_col, value=usd_header)
                copy_style(src_head, hc)
                hc.font = font_bold
            elif isinstance(src_head.value, str) and src_head.value.strip() and hr <= max(header_row, 2):
                hc = ws.cell(row=hr, column=usd_col, value=src_head.value)
                copy_style(src_head, hc)

        max_rs_len = len(header_text)
        max_usd_len = len(usd_header) + 2

        for row in range(header_row + 1, ws.max_row + 1):
            rs_val = ws_values.cell(row=row, column=orig_col).value
            if rs_val is None or isinstance(rs_val, bool) or not isinstance(rs_val, (int, float)):
                continue
            formula = f"={data_letter}{row}/{rate_ref}"
            cell = ws.cell(row=row, column=usd_col, value=formula)
            cell.number_format = num_fmt
            src_font = ws.cell(row=row, column=data_col).font
            cell.font = Font(name="Arial", size=8, bold=src_font.b, color="FF323232")

            usd_val = float(rs_val) / rate
            max_rs_len = max(max_rs_len, len(f"{rs_val:,.2f}"))
            max_usd_len = max(max_usd_len, len(f"${usd_val:,.2f}"))

            if is_grand_total:
                label = find_row_label(ws, row)
                summary.append({"label": label, "row": row, "rs": float(rs_val), "usd": usd_val})

        ws.column_dimensions[data_letter].width = min(max(max_rs_len * 1.2 + 3, 10), 60)
        ws.column_dimensions[get_column_letter(usd_col)].width = min(max(max_usd_len * 1.2 + 3, 10), 60)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read(), summary
