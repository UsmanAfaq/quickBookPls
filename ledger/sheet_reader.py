"""Reading a QuickBooks P&L export: finding the TOTAL column, company/group
columns, row labels, outline levels. Nothing here writes Excel/PDF."""
import io

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

TOTAL_HEADER_TEXT = "total"
HEADER_SCAN_ROWS = 5
LABEL_SCAN_COLS = 8

# Excel export: rows deeper than this outline level start collapsed.
OUTLINE_COLLAPSE_FROM_LEVEL = 2
# Excel export: the first label columns (A:E) are made narrower.
LABEL_NARROW_COLS = 5
LABEL_NARROW_FACTOR = 5
LABEL_NARROW_MIN = 2.0

PROMINENT_LABELS = {"gross profit", "net ordinary income", "net income"}
PNL_TOTAL_PREFIXES = (
    "total ",
    "gross profit",
    "net ordinary income",
    "net other income",
    "net income",
)


def norm_label(label):
    """Lower-case + collapse repeated spaces."""
    return " ".join(str(label).lower().split())


def find_total_column(ws):
    for row in range(1, HEADER_SCAN_ROWS + 1):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            if isinstance(val, str) and val.strip().lower() == TOTAL_HEADER_TEXT:
                return col, row
    return _find_unlabelled_total_column(ws)


def _find_unlabelled_total_column(ws):
    """Some exports (e.g. Oct-26) have the grand-total column but no "Total" header cell.
    It is the column right after the last labelled company/group column and it holds numbers.
    Detecting it means the file never has to be rewritten (a re-save would drop the cached
    results of its formulas)."""
    hr = _find_group_header_row(ws)
    if not hr:
        return None, None
    last_labelled = 0
    for c in range(2, ws.max_column + 1):
        if any(isinstance(ws.cell(row=r, column=c).value, str) and ws.cell(row=r, column=c).value.strip()
               for r in range(1, hr + 1)):
            last_labelled = c
    cand = last_labelled + 1
    if not last_labelled or cand > ws.max_column:
        return None, None
    for r in range(hr + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=cand).value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return cand, hr
        if isinstance(v, str) and v.startswith("="):
            return cand, hr
    return None, None


def column_header_info(ws, col, header_row):
    """(name, name_row) of a data column. Row 1 wins, then row 2, then header_row."""
    g = _find_group_header_row(ws) or 2
    for r in (g - 1, g, header_row):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and v.strip() and not v.strip().startswith("="):
            return v.strip(), r
    return "", header_row


def find_row_label(ws, row):
    """Nearest text label (rightmost in the label columns) on this row."""
    label = None
    for col in range(1, LABEL_SCAN_COLS + 1):
        val = ws.cell(row=row, column=col).value
        if isinstance(val, str) and val.strip() and not val.strip().startswith("="):
            label = val.strip()
    return label or f"Row {row}"


def find_row_label_col(ws, row):
    """Column number of the first text label on this row."""
    for col in range(1, LABEL_SCAN_COLS + 1):
        val = ws.cell(row=row, column=col).value
        if isinstance(val, str) and val.strip() and not val.strip().startswith("="):
            return col
    return None


def _find_group_header_row(ws):
    """Row (within the header scan area) holding the "(group)" / "Total ..." labels."""
    for row in range(1, HEADER_SCAN_ROWS + 1):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            if isinstance(val, str):
                v = val.strip().lower()
                if v.startswith("total ") or (v.startswith("(") and v.endswith(")")):
                    return row
    return None


def ensure_total_column(file_bytes: bytes) -> bytes:
    """If the export has no TOTAL column, append one that sums every company column
    (group "Total ..." columns are skipped so nothing is counted twice).
    Returns the original bytes untouched when a TOTAL header already exists."""
    wb = load_workbook(io.BytesIO(file_bytes))
    wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True)
    for sheet in wb.worksheets:
        if find_total_column(sheet)[0]:
            return file_bytes
    for ws, wsv in zip(wb.worksheets, wb_values.worksheets):
        header_row = _find_group_header_row(ws)
        if not header_row:
            continue
        last_col = 0
        for c in range(1, ws.max_column + 1):
            if any(isinstance(ws.cell(row=r, column=c).value, str)
                   and ws.cell(row=r, column=c).value.strip()
                   for r in range(1, header_row + 1)):
                last_col = c
        if not last_col:
            continue
        sum_cols = []
        for c in range(1, last_col + 1):
            hdr = ws.cell(row=header_row, column=c).value
            has_name = any(isinstance(ws.cell(row=r, column=c).value, str)
                           and ws.cell(row=r, column=c).value.strip()
                           for r in range(1, header_row + 1))
            if not has_name or c <= 1:
                continue
            if isinstance(hdr, str) and hdr.strip().lower().startswith("total "):
                continue
            sum_cols.append(c)
        total_col = last_col + 1
        for r in range(1, wsv.max_row + 1):
            if r == header_row:
                ws.cell(row=r, column=total_col).value = "TOTAL"
            elif r > header_row:
                nums = [wsv.cell(row=r, column=c).value for c in sum_cols]
                nums = [v for v in nums if isinstance(v, (int, float)) and not isinstance(v, bool)]
                if nums:
                    ws.cell(row=r, column=total_col).value = sum(nums)
            else:
                continue
            src, dst = ws.cell(row=r, column=last_col), ws.cell(row=r, column=total_col)
            if src.has_style:
                dst._style = src._style
        letter = get_column_letter(last_col)
        width = ws.column_dimensions[letter].width
        if width:
            ws.column_dimensions[get_column_letter(total_col)].width = width
        # openpyxl drops the cached result of every formula on save, so formula cells
        # (e.g. a group's "Total ..." column =SUM(B:F)) would read back as empty / 0.
        # Freeze each formula to the value Excel had calculated; the app only reads values.
        for row in wsv.iter_rows():
            for vc in row:
                fc = ws.cell(row=vc.row, column=vc.column)
                if (isinstance(fc.value, str) and fc.value.startswith("=")
                        and vc.value is not None):
                    fc.value = vc.value
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()
    return file_bytes


def load_sheet(file_bytes: bytes):
    """Returns (ws_formulas, ws_values, total_col, header_row)."""
    wb = load_workbook(io.BytesIO(file_bytes))
    wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True)
    for sheet, sheet_values in zip(wb.worksheets, wb_values.worksheets):
        col, row = find_total_column(sheet)
        if col:
            return sheet, sheet_values, col, row
    raise ValueError('Could not find a "TOTAL" header on any sheet.')


def load_values_sheet(file_bytes: bytes):
    """The data_only worksheet that has the TOTAL header."""
    wb_values = load_workbook(io.BytesIO(file_bytes), data_only=True)
    for sheet_values in wb_values.worksheets:
        col, row = find_total_column(sheet_values)
        if col:
            return sheet_values
    raise ValueError('Could not find a "TOTAL" header on any sheet.')


def sum_cell_across_files(ws_values_list, row, col):
    """Returns (total, found) for one cell summed across several months."""
    total = 0.0
    found = False
    for wsv in ws_values_list:
        if row > wsv.max_row:
            continue
        val = wsv.cell(row=row, column=col).value
        if val is None or isinstance(val, bool) or not isinstance(val, (int, float)):
            continue
        total += float(val)
        found = True
    return total, found


def find_companies(ws, total_col, header_row):
    companies = []
    seen = set()
    g = _find_group_header_row(ws) or 2   # exports may have extra rows (e.g. a date) on top
    for col in range(1, total_col):
        r1 = ws.cell(row=g - 1, column=col).value
        r2 = ws.cell(row=g, column=col).value
        name = None
        if isinstance(r1, str) and r1.strip():
            name = r1.strip()
        elif isinstance(r2, str) and r2.strip():
            name = r2.strip()
        if name and name not in seen:
            seen.add(name)
            companies.append({"name": name, "col": col})
    return companies


def find_groups_and_standalone(ws, total_col, last_col=None):
    """Returns (groups, standalone). groups = {name: {"name","members","total_col"}}.
    `last_col` > total_col: extra columns placed after TOTAL (see build_reference)."""
    groups = {}
    standalone = []
    g = _find_group_header_row(ws) or 2
    cols = list(range(1, total_col)) + list(range((total_col + 1), (last_col or 0) + 1))
    for c in cols:
        r1 = ws.cell(row=g - 1, column=c).value
        r2 = ws.cell(row=g, column=c).value
        r1 = r1.strip() if isinstance(r1, str) and r1.strip() else None
        r2 = r2.strip() if isinstance(r2, str) and r2.strip() else None
        if r2 and r2.startswith("(") and r2.endswith(")"):
            gname = r2[1:-1].strip()
            groups.setdefault(gname, {"name": gname, "members": [], "total_col": None})
            groups[gname]["members"].append({"name": r1 or f"Column {c}", "col": c})
        elif r2 and r2.lower().startswith("total "):
            gname = r2[6:].strip()
            groups.setdefault(gname, {"name": gname, "members": [], "total_col": None})
            groups[gname]["total_col"] = c
        elif r2 and not r1:
            standalone.append({"name": r2, "col": c})
        elif r1 and not r2:
            standalone.append({"name": r1, "col": c})
    return groups, standalone


_TOTAL_ALIASES = {"cost of goods sold": "cogs"}
_TOP_SECTIONS = {"ordinary income/expense", "other income/expense"}


def _flat_label_columns(ws, label_cols):
    """Virtual label column per row for a flat sheet, so the normal outline logic applies.
    A row with no numbers is a section header; it opens a group that its matching
    "Total <name>" row closes (the total sits at the header's own level)."""
    def has_data(r):
        return any(ws.cell(row=r, column=c).value is not None
                   for c in range(2, ws.max_column + 1))

    rows = sorted(label_cols)
    names = {r: norm_label(ws.cell(row=r, column=label_cols[r]).value) for r in rows}
    alias = lambda n: _TOTAL_ALIASES.get(n, n)
    closing = {alias(n[6:]) for n in names.values() if n.startswith("total ")}

    virtual, stack = {}, []          # stack: [(alias name, virtual col)]
    for r in rows:
        n = names[r]
        if n in _TOP_SECTIONS:
            stack = [(n, 1)]
            virtual[r] = 1
            continue
        if n.startswith("total "):
            key = alias(n[6:])
            hit = next((i for i in range(len(stack) - 1, -1, -1) if stack[i][0] == key), None)
            if hit is not None:
                virtual[r] = stack[hit][1]
                del stack[hit:]
                continue
        if n.startswith(PNL_TOTAL_PREFIXES[1:]):      # Gross Profit / Net ... Income
            del stack[1:]
            virtual[r] = 2
            continue
        virtual[r] = len(stack) + 1
        if not has_data(r) and alias(n) in closing:
            stack.append((alias(n), virtual[r]))
    return virtual


def compute_row_outline(ws, first_row, last_row):
    """Excel outline levels from label indentation. Returns (levels, closers)."""
    label_cols = {}
    for r in range(first_row, last_row + 1):
        c = find_row_label_col(ws, r)
        if c:
            label_cols[r] = c
    if len(set(label_cols.values())) == 1:
        # Flat export (every label in one column, no indentation): rebuild the
        # hierarchy from section headers and their "Total <section>" rows.
        label_cols = _flat_label_columns(ws, label_cols)
    rows = sorted(label_cols)
    levels, closers, stack = {}, {}, []
    for i, r in enumerate(rows):
        c = label_cols[r]
        while stack and stack[-1][0] >= c:
            col_h, _h_row, h_level = stack.pop()
            if col_h == c and r not in closers:
                closers[r] = h_level + 1
        level = len(stack)
        levels[r] = level
        nxt = label_cols[rows[i + 1]] if i + 1 < len(rows) else None
        if nxt is not None and nxt > c:
            stack.append((c, r, level))
    return levels, closers



# ---- per-month column mapping ---------------------------------------------
# Every monthly export can have a different layout (extra columns, shifted TOTAL
# column, different header row, different row count), so a column picked on the
# first file must be re-located in each other file by its header text.

def _header_text(ws, row, col):
    v = ws.cell(row=row, column=col).value
    return norm_label(v) if isinstance(v, str) and v.strip() else None


def column_key(ws, col):
    """(company name, group tag) of a column, whitespace/case-normalised."""
    g = _find_group_header_row(ws) or 2
    return _header_text(ws, g - 1, col), _header_text(ws, g, col)


def map_column(ws_src, col, total_src, ws_dst, total_dst):
    """Column in ws_dst matching `col` of ws_src, or None when it does not exist there."""
    if col == total_src:
        return total_dst
    key = column_key(ws_src, col)
    if key == (None, None):
        return None
    candidates = [(c, column_key(ws_dst, c)) for c in range(1, total_dst)]
    for c, k in candidates:
        if k == key:
            return c
    # a group's "Total ..." column is identified by that label alone (its company-name
    # cell may differ between months, e.g. "3. OZI GameMax - Other" in one sheet only)
    if key[1] and key[1].startswith("total "):
        for c, k in candidates:
            if k[1] == key[1]:
                return c
    # fall back to the entity name alone (its group tag may have been renamed)
    name = key[0] or key[1]
    for c, k in candidates:
        if name and name in (k[0] or (None if k[1] is None else k[1]),):
            return c
    return None

# ---- reference layout = first sheet + every column that only appears in later sheets
def build_reference(contents):
    """(ws, total_col, header_row, last_col) for a set of monthly files.

    Monthly exports gain / lose teams over time (e.g. Melon Games and OZI South only
    appear from Feb / Jun). Using only the first sheet's columns hides those teams, so
    this returns a small in-memory sheet that keeps the FIRST file's column numbers
    exactly, and appends every team that only exists in later files AFTER its TOTAL
    column (total_col + 1 ... last_col). Per-month values are located later by header
    text (map_column), so these numbers are just ids for the dashboard."""
    from openpyxl import Workbook
    sheets = []
    for b in contents:
        wsv = load_values_sheet(b)
        sheets.append((wsv, find_total_column(wsv)[0]))
    base, tc = sheets[0]
    hr = find_total_column(base)[1]
    g = _find_group_header_row(base) or 2
    ref = Workbook().active
    for r in range(1, HEADER_SCAN_ROWS + 1):
        for c in range(1, tc + 1):
            v = base.cell(row=r, column=c).value
            if v is not None:
                ref.cell(row=r, column=c).value = v
    last = tc
    for wsv, tcv in sheets[1:]:
        gv = _find_group_header_row(wsv) or 2
        for c in range(2, tcv):
            if column_key(wsv, c) == (None, None):
                continue
            if map_column(wsv, c, tcv, ref, last + 1) is not None:
                continue
            last += 1
            for src_row, dst_row in ((gv - 1, g - 1), (gv, g)):
                v = wsv.cell(row=src_row, column=c).value
                if isinstance(v, str) and v.strip():
                    ref.cell(row=dst_row, column=last).value = v.strip()
    return ref, tc, hr, last