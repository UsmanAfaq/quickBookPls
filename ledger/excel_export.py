"""Filtered Excel download (selected companies x months, optional comparison blocks)."""
import io
import math
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import Outline

from .rates import as_rates, month_pairs
from .sheet_reader import (LABEL_NARROW_COLS, LABEL_NARROW_FACTOR, LABEL_NARROW_MIN,
                           OUTLINE_COLLAPSE_FROM_LEVEL, PNL_TOTAL_PREFIXES, PROMINENT_LABELS,
                           column_header_info, compute_row_outline, find_groups_and_standalone,
                           find_row_label_col, find_total_column, load_values_sheet,
                           norm_label, sum_cell_across_files)
from .usd_column import copy_style

PKR_FMT_NEG_RED = '#,##0.00;[Red]-#,##0.00'
USD_FMT_NEG_RED = '"$"#,##0.00;[Red]-"$"#,##0.00'

XL_COLORS = {
    "month_bg": "FF1E293B",
    "pkr_head": "FF2F5D8C",
    "pkr_cell": "FFEAF1F8",
    "pkr_key": "FFC9DCEE",
    "usd_head": "FF5B84B1",
    "usd_cell": "FFF5F9FC",
    "usd_key": "FFDDE9F5",
    "tot_head": "FF334155",
    "tot_cell": "FFF1F5F9",
    "tot_key": "FFCBD5E1",
    "label_key": "FFE2E8F0",
    "label_total": "FFF1F5F9",
    "label_head": "FFF8FAFC",
    "light_line": "FFE2E8F0",
    "dark_line": "FF475569",
}

# (banner, PKR header, USD header, PKR body, USD body, PKR key-row, USD key-row)
# Muted, same-saturation hues: blue, teal, indigo, sand, rose, slate.
MONTH_PALETTES = [
    ("FF1E3A5F", "FF2F5D8C", "FF5B84B1", "FFEAF1F8", "FFF5F9FC", "FFC9DCEE", "FFDDE9F5"),
    ("FF134E4A", "FF1F766E", "FF4F9D96", "FFE6F4F2", "FFF3FAF9", "FFBFE3DE", "FFD7EDEA"),
    ("FF3B3A75", "FF54539A", "FF8281BD", "FFEDEDF8", "FFF6F6FC", "FFD3D2EC", "FFE4E4F4"),
    ("FF6B4A1E", "FF8C6A34", "FFB89458", "FFF6EFE3", "FFFBF7F0", "FFE6D6B8", "FFF1E6D0"),
    ("FF7A2E48", "FF9E4561", "FFC27B90", "FFF8E9EE", "FFFCF4F7", "FFEBC9D5", "FFF3DCE4"),
    ("FF334155", "FF475569", "FF7B8CA3", "FFEDF0F4", "FFF6F8FA", "FFD0D7E1", "FFE1E6ED"),
]


def _fill(hex_argb):
    return PatternFill(fill_type="solid", start_color=hex_argb, end_color=hex_argb)


def combined_columns(ws, total_col, cols):
    """Which selected columns can be added into the Combined Total without
    double counting. Returns (use, skipped)."""
    colset = set(cols)
    if total_col in colset:
        return [total_col], [c for c in cols if c != total_col]
    groups, _standalone = find_groups_and_standalone(ws, total_col)
    skip = set()
    for g in groups.values():
        if g["total_col"] in colset:
            skip |= {m["col"] for m in g["members"]}
    return [c for c in cols if c not in skip], [c for c in cols if c in skip]


def safe_sheet_title(name, n_selected=1):
    """Excel sheet names: max 31 chars, none of  \\ / ? * [ ] :"""
    clean = "".join("-" if ch in '\\/?*[]:' else ch for ch in (name or "")).strip()
    if len(clean) <= 31 and clean:
        return clean
    if n_selected > 1:
        return f"Export ({n_selected} selected)"
    return (clean[:31] or "Export")


def export_selection(files_bytes, rate, cols, sheet_title: str, month_names=None, compare=False):
    """Builds a new workbook with label columns + the given data columns.
    Layout: label columns | MONTH 1 block | MONTH 2 block | ... | Combined Total
    compare=True (2+ months): one "Comparison in Amount" block per pair of months."""
    if not isinstance(files_bytes, (list, tuple)):
        files_bytes = [files_bytes]
    if not month_names or len(month_names) != len(files_bytes):
        month_names = [f"File {i + 1}" for i in range(len(files_bytes))]

    wb = load_workbook(io.BytesIO(files_bytes[0]))
    ws = None
    total_col = header_row = None
    for sheet in wb.worksheets:
        col, row = find_total_column(sheet)
        if col:
            ws, total_col, header_row = sheet, col, row
            break
    if ws is None:
        raise ValueError('Could not find a "TOTAL" header on any sheet.')

    ws_values_list = [load_values_sheet(fb) for fb in files_bytes]
    rates = as_rates(rate, len(files_bytes))

    groups_, standalone_ = find_groups_and_standalone(ws, total_col)
    all_data_cols = sorted(m["col"] for g in groups_.values() for m in g["members"]) \
        + [c["col"] for c in standalone_]
    label_col_count = (min(all_data_cols) - 1) if all_data_cols else 7

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = safe_sheet_title(sheet_title, len(cols))
    out_ws.sheet_view.showGridLines = False

    RO = 1
    max_row = ws.max_row
    last_out_row = max_row + RO
    head_rows = max(header_row, 2)
    use_cols, skipped_cols = combined_columns(ws, total_col, cols)
    colset_sel = set(cols)
    compare = bool(compare) and len(files_bytes) >= 2
    add_combined = (len(cols) >= 2 or len(files_bytes) > 1) and bool(use_cols) and not compare

    light = Side(style="thin", color=XL_COLORS["light_line"])
    light_border = Border(left=light, right=light, top=light, bottom=light)

    for r in range(1, max_row + 1):
        h = ws.row_dimensions[r].height
        if h is not None:
            out_ws.row_dimensions[r + RO].height = h
    out_ws.row_dimensions[1].height = 24

    # ---- which source rows are headings / totals / key result rows ---------
    row_kind = {}
    for r in range(header_row + 1, max_row + 1):
        lc = find_row_label_col(ws, r)
        if not lc:
            continue
        low = norm_label(ws.cell(row=r, column=lc).value)
        has_num = False
        for wsv in ws_values_list:
            if r > wsv.max_row:
                continue
            for c in cols:
                v = wsv.cell(row=r, column=c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    has_num = True
                    break
            if has_num:
                break
        if low in PROMINENT_LABELS:
            row_kind[r] = "key"
        elif not has_num:
            row_kind[r] = "heading"
        elif low.startswith(PNL_TOTAL_PREFIXES):
            row_kind[r] = "total"

    # ---- label / description columns --------------------------------------
    for lc in range(1, label_col_count + 1):
        for r in range(1, max_row + 1):
            src = ws.cell(row=r, column=lc)
            cell = out_ws.cell(row=r + RO, column=lc)
            if src.value is not None:
                cell.value = src.value
            copy_style(src, cell)
            cell.border = light_border
            kind = row_kind.get(r)
            if kind in ("heading", "total", "key"):
                f = copy(cell.font)
                cell.font = Font(name=f.name or "Arial", size=(10 if kind == "key" else (f.sz or 9)),
                                 bold=True, color="FF000000")
            if kind == "key":
                cell.fill = _fill(XL_COLORS["label_key"])
            elif kind == "total":
                cell.fill = _fill(XL_COLORS["label_total"])
            elif kind == "heading":
                cell.fill = _fill(XL_COLORS["label_head"])
            else:
                cell.fill = PatternFill(fill_type=None)
        src_letter = get_column_letter(lc)
        if src_letter in ws.column_dimensions and ws.column_dimensions[src_letter].width:
            out_ws.column_dimensions[get_column_letter(lc)].width = ws.column_dimensions[src_letter].width
        if lc <= LABEL_NARROW_COLS:
            _w = out_ws.column_dimensions[get_column_letter(lc)].width or 8.43
            out_ws.column_dimensions[get_column_letter(lc)].width = max(round(_w * LABEL_NARROW_FACTOR, 2), LABEL_NARROW_MIN)
        banner = out_ws.cell(row=1, column=lc)
        banner.fill = _fill(XL_COLORS["month_bg"])
    out_ws.cell(row=1, column=1, value="P&L — Rs / USD").font = Font(name="Arial", size=11, bold=True, color="FFFFFFFF")

    # ---- data blocks: MONTH -> company -> (PKR, USD) -----------------------
    out_col = label_col_count + 1
    positions = {}
    company_boxes = []
    month_boxes = []
    pkr_cols, usd_cols, tot_cols = [], [], []

    def head_len(cpos):
        return max((len(str(out_ws.cell(row=hr + RO, column=cpos).value))
                    for hr in range(1, head_rows + 1)
                    if isinstance(out_ws.cell(row=hr + RO, column=cpos).value, str)), default=0)

    def style_header(cpos, kind, pal=None):
        if kind == "tot" or pal is None:
            fill = _fill(XL_COLORS["tot_head"])
        else:
            fill = _fill(pal[1] if kind == "pkr" else pal[2])
        for hr in range(1, head_rows + 1):
            c = out_ws.cell(row=hr + RO, column=cpos)
            c.fill = fill
            c.font = Font(name="Arial", size=9, bold=True, color="FFFFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ---- layout: a company/group with 2+ selected teams gets a TOTAL column
    # (always visible) with its teams tucked under an Excel "+" (collapsed) ----
    member_group = {m["col"]: g["name"] for g in groups_.values() for m in g["members"]}
    sel_members = {}
    for c in cols:
        if c in member_group:
            sel_members.setdefault(member_group[c], []).append(c)
    ordered = []
    done_groups = set()
    group_total_cols = {g["total_col"]: g["name"] for g in groups_.values()
                        if g["total_col"] and len(sel_members.get(g["name"], [])) >= 2}
    for c in cols:
        if c in group_total_cols:
            gname = group_total_cols[c]
            ordered.append(("col", c, 0))
            for m in sel_members[gname]:
                ordered.append(("col", m, 1))
            done_groups.add(gname)
        elif c in member_group:
            gname = member_group[c]
            if gname in done_groups:
                continue
            members = sel_members[gname]
            if len(members) >= 2:
                if groups_[gname]["total_col"] in colset_sel:
                    continue  # emitted together with its group total column
                ordered.append(("sumtotal", gname, members))
                for m in members:
                    ordered.append(("col", m, 1))
                done_groups.add(gname)
            else:
                ordered.append(("col", c, 0))
        else:
            ordered.append(("col", c, 0))

    for month_idx, wsv in enumerate(ws_values_list):
        pal = MONTH_PALETTES[month_idx % len(MONTH_PALETTES)]
        rate = rates[month_idx]
        rate_lit = f"{rate:.10g}"
        month_start = out_col
        for item in ordered:
            if item[0] == "sumtotal":
                gname, members = item[1], item[2]
                t_pkr, t_usd = out_col, out_col + 1
                first_name, t_name_row = column_header_info(ws, members[0], header_row)
                out_ws.cell(row=t_name_row + RO, column=t_pkr, value=f"{gname} — TOTAL ({len(members)} teams) (PKR)")
                out_ws.cell(row=t_name_row + RO, column=t_usd, value=f"{gname} — TOTAL ({len(members)} teams) (USD)")
                style_header(t_pkr, "tot")
                style_header(t_usd, "tot")
                _hr = out_ws.row_dimensions[t_name_row + RO]
                _hr.height = max(_hr.height or 0, 40)
                mp = [out_col + 2 + 2 * i for i in range(len(members))]
                t_len = 0
                for r in range(header_row + 1, max_row + 1):
                    orow = r + RO
                    kind = row_kind.get(r)
                    key = kind == "key"
                    font = Font(name="Arial", size=10 if key else 9,
                                bold=key or kind in ("heading", "total"), color="FF000000")
                    for cp in (t_pkr, t_usd):
                        cell = out_ws.cell(row=orow, column=cp)
                        cell.fill = _fill(XL_COLORS["tot_key" if key else "tot_cell"])
                        cell.font = font
                        cell.alignment = Alignment(horizontal="right")
                    tv, found = 0.0, False
                    for c in members:
                        v = wsv.cell(row=r, column=c).value if r <= wsv.max_row else None
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            tv += float(v)
                            found = True
                    if not found:
                        continue
                    out_ws.cell(row=orow, column=t_pkr,
                                value="=" + "+".join(f"{get_column_letter(p)}{orow}" for p in mp)).number_format = PKR_FMT_NEG_RED
                    out_ws.cell(row=orow, column=t_usd,
                                value="=" + "+".join(f"{get_column_letter(p + 1)}{orow}" for p in mp)).number_format = USD_FMT_NEG_RED
                    t_len = max(t_len, len(f"{tv:,.2f}"), len(f"{tv / rate:,.2f}") + 1)
                w = min(max(24, head_len(t_pkr) * 0.6 + 4, t_len + 5), 60)
                out_ws.column_dimensions[get_column_letter(t_pkr)].width = w
                out_ws.column_dimensions[get_column_letter(t_usd)].width = w
                company_boxes.append((t_pkr, t_usd))
                out_col += 2
                continue
            col, outline = item[1], item[2]
            header_text, name_row = column_header_info(ws, col, header_row)
            pkr_pos, usd_pos = out_col, out_col + 1
            positions[(col, month_idx)] = pkr_pos
            pkr_cols.append(pkr_pos)
            usd_cols.append(usd_pos)
            company_boxes.append((pkr_pos, usd_pos))

            for r in range(1, head_rows + 1):
                src = ws.cell(row=r, column=col)
                is_text = isinstance(src.value, str) and src.value.strip()
                if src.value is not None and (r <= header_row or is_text):
                    out_ws.cell(row=r + RO, column=pkr_pos, value=src.value)
                if r != name_row and is_text:
                    out_ws.cell(row=r + RO, column=usd_pos, value=src.value)
            out_ws.cell(row=name_row + RO, column=pkr_pos,
                        value=f"{header_text} (PKR)" if header_text else "PKR")
            out_ws.cell(row=name_row + RO, column=usd_pos,
                        value=f"{header_text} (USD)" if header_text else "USD")
            style_header(pkr_pos, "pkr", pal)
            style_header(usd_pos, "usd", pal)

            data_letter = get_column_letter(pkr_pos)
            for r in range(header_row + 1, max_row + 1):
                src = ws.cell(row=r, column=col)
                orow = r + RO
                kind = row_kind.get(r)
                key = kind == "key"
                bold = key or kind in ("heading", "total") or bool(src.font and src.font.b)
                rs_val = wsv.cell(row=r, column=col).value if r <= wsv.max_row else None

                value = None
                if rs_val is not None and not isinstance(rs_val, bool) and isinstance(rs_val, (int, float)):
                    value = float(rs_val)
                elif isinstance(src.value, str) and src.value.startswith("="):
                    try:
                        value = Translator(src.value, origin=src.coordinate).translate_formula(
                            out_ws.cell(row=orow, column=pkr_pos).coordinate)
                    except Exception:
                        value = src.value

                pc = out_ws.cell(row=orow, column=pkr_pos)
                uc = out_ws.cell(row=orow, column=usd_pos)
                font = Font(name="Arial", size=10 if key else 9, bold=bold, color="FF000000")
                pc.fill = _fill(pal[5] if key else pal[3])
                uc.fill = _fill(pal[6] if key else pal[4])
                pc.font = font
                uc.font = font
                pc.alignment = Alignment(horizontal="right")
                uc.alignment = Alignment(horizontal="right")
                if value is not None:
                    pc.value = value
                    pc.number_format = PKR_FMT_NEG_RED
                    uc.value = f"={data_letter}{orow}/{rate_lit}"
                    uc.number_format = USD_FMT_NEG_RED

            base_w = 10
            sl = get_column_letter(col)
            if sl in ws.column_dimensions and ws.column_dimensions[sl].width:
                base_w = ws.column_dimensions[sl].width
            num_len = 0
            for rr in range(header_row + 1, min(max_row, wsv.max_row) + 1):
                vv = wsv.cell(row=rr, column=col).value
                if isinstance(vv, (int, float)) and not isinstance(vv, bool):
                    num_len = max(num_len, len(f"{vv:,.2f}"), len(f"{vv / rate:,.2f}") + 1)
            out_ws.column_dimensions[get_column_letter(pkr_pos)].width = min(max(base_w, head_len(pkr_pos) * 1.1 + 2, num_len + 5), 60)
            out_ws.column_dimensions[get_column_letter(usd_pos)].width = min(max(base_w, head_len(usd_pos) * 1.1 + 2, num_len + 5), 60)
            if outline:
                for cp in (pkr_pos, usd_pos):
                    d = out_ws.column_dimensions[get_column_letter(cp)]
                    d.outline_level = 1   # still groupable (click "-" to collapse), but shown by default
            out_col += 2

        month_boxes.append((month_start, out_col - 1))
        for c in range(month_start, out_col):
            cell = out_ws.cell(row=1, column=c)
            cell.fill = _fill(pal[0])
            cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        out_ws.cell(row=1, column=month_start, value=month_names[month_idx])
        if out_col - 1 > month_start:
            out_ws.merge_cells(start_row=1, start_column=month_start, end_row=1, end_column=out_col - 1)

    # ---- COMPARISON IN AMOUNT blocks --------------------------------------
    if compare:
        steps = [{"a": a, "b": b, "overall": ov, "title": t}
                 for a, b, t, ov in month_pairs(month_names)]
        pct_fmt = '0.00%;[Red]-0.00%'
        for st in steps:
            if st["overall"]:
                pal_h, pal_b, pal_k, pal_ban = XL_COLORS["tot_head"], XL_COLORS["tot_cell"], XL_COLORS["tot_key"], XL_COLORS["month_bg"]
            else:
                pal_h, pal_b, pal_k, pal_ban = "FF475569", "FFF4F6F9", "FFD7DDE6", "FF334155"
            blk_start = out_col
            diff_cols = {}
            for col in cols:
                header_text, name_row = column_header_info(ws, col, header_row)
                c_pkr, c_pct, c_usd = out_col, out_col + 1, out_col + 2
                diff_cols[col] = (c_pkr, c_pct, c_usd)
                labels3 = (f"{header_text} — Difference (PKR)" if header_text else "Difference (PKR)",
                           f"{header_text} — Difference %" if header_text else "Difference %",
                           f"{header_text} — Difference (USD)" if header_text else "Difference (USD)")
                for r in range(1, head_rows + 1):
                    src = ws.cell(row=r, column=col)
                    if isinstance(src.value, str) and src.value.strip() and r != name_row:
                        for cc in (c_pkr, c_pct, c_usd):
                            out_ws.cell(row=r + RO, column=cc, value=src.value)
                for cc, lab in zip((c_pkr, c_pct, c_usd), labels3):
                    out_ws.cell(row=name_row + RO, column=cc, value=lab)
                    for hr in range(1, head_rows + 1):
                        h = out_ws.cell(row=hr + RO, column=cc)
                        h.fill = _fill(pal_h)
                        h.font = Font(name="Arial", size=9, bold=True, color="FFFFFFFF")
                        h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                _hr = out_ws.row_dimensions[name_row + RO]
                _hr.height = max(_hr.height or 0, 40)

                pa, pb = positions[(col, st["a"])], positions[(col, st["b"])]
                for r in range(header_row + 1, max_row + 1):
                    orow = r + RO
                    kind = row_kind.get(r)
                    key = kind == "key"
                    font = Font(name="Arial", size=10 if key else 9,
                                bold=key or kind in ("heading", "total"), color="FF000000")
                    cells = [out_ws.cell(row=orow, column=cc) for cc in (c_pkr, c_pct, c_usd)]
                    for cell in cells:
                        cell.fill = _fill(pal_k if key else pal_b)
                        cell.font = font
                        cell.alignment = Alignment(horizontal="right")
                    va = out_ws.cell(row=orow, column=pa).value
                    vb = out_ws.cell(row=orow, column=pb).value
                    is_num = lambda v: (isinstance(v, (int, float)) and not isinstance(v, bool)) or (isinstance(v, str) and v.startswith("="))
                    if not (is_num(va) or is_num(vb)):
                        continue
                    ea, eb = f"{get_column_letter(pa)}{orow}", f"{get_column_letter(pb)}{orow}"
                    ua, ub = f"{get_column_letter(pa + 1)}{orow}", f"{get_column_letter(pb + 1)}{orow}"
                    d_ref = f"{get_column_letter(c_pkr)}{orow}"
                    cells[0].value = f"={eb}-{ea}"
                    cells[0].number_format = PKR_FMT_NEG_RED
                    cells[1].value = f'=IF({ea}=0,"",{d_ref}/ABS({ea}))'
                    cells[1].number_format = pct_fmt
                    cells[2].value = f"={ub}-{ua}"
                    cells[2].number_format = USD_FMT_NEG_RED
                out_ws.column_dimensions[get_column_letter(c_pkr)].width = 20
                out_ws.column_dimensions[get_column_letter(c_pct)].width = 14
                out_ws.column_dimensions[get_column_letter(c_usd)].width = 20
                company_boxes.append((c_pkr, c_usd))
                out_col += 3

            # ---- TOTAL of all selected companies for this comparison step ----
            tot_members = [c for c in use_cols if c in diff_cols]
            if len(tot_members) >= 1:
                t_pkr, t_pct, t_usd = out_col, out_col + 1, out_col + 2
                _n, t_name_row = column_header_info(ws, cols[0], header_row)
                tag = "TOTAL (all selected)" if len(tot_members) > 1 else "TOTAL"
                for cc, lab in zip((t_pkr, t_pct, t_usd),
                                   (f"{tag} — Difference (PKR)",
                                    f"{tag} — Difference %",
                                    f"{tag} — Difference (USD)")):
                    out_ws.cell(row=t_name_row + RO, column=cc, value=lab)
                    for hr in range(1, head_rows + 1):
                        h = out_ws.cell(row=hr + RO, column=cc)
                        h.fill = _fill(XL_COLORS["tot_head"])
                        h.font = Font(name="Arial", size=9, bold=True, color="FFFFFFFF")
                        h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                _hr = out_ws.row_dimensions[t_name_row + RO]
                _hr.height = max(_hr.height or 0, 40)

                for r in range(header_row + 1, max_row + 1):
                    orow = r + RO
                    kind = row_kind.get(r)
                    key = kind == "key"
                    font = Font(name="Arial", size=10 if key else 9,
                                bold=key or kind in ("heading", "total"), color="FF000000")
                    cells = [out_ws.cell(row=orow, column=cc) for cc in (t_pkr, t_pct, t_usd)]
                    for cell in cells:
                        cell.fill = _fill(XL_COLORS["tot_key"] if key else XL_COLORS["tot_cell"])
                        cell.font = font
                        cell.alignment = Alignment(horizontal="right")

                    pkr_refs, usd_refs, base_refs = [], [], []
                    for c in tot_members:
                        d_pkr, _d_pct, d_usd = diff_cols[c]
                        if out_ws.cell(row=orow, column=d_pkr).value is None:
                            continue
                        pkr_refs.append(f"{get_column_letter(d_pkr)}{orow}")
                        usd_refs.append(f"{get_column_letter(d_usd)}{orow}")
                        base_refs.append(f"{get_column_letter(positions[(c, st['a'])])}{orow}")
                    if not pkr_refs:
                        continue
                    cells[0].value = "=" + "+".join(pkr_refs)
                    cells[0].number_format = PKR_FMT_NEG_RED
                    base = "(" + "+".join(base_refs) + ")"
                    cells[1].value = f'=IF({base}=0,"",{get_column_letter(t_pkr)}{orow}/ABS({base}))'
                    cells[1].number_format = pct_fmt
                    cells[2].value = "=" + "+".join(usd_refs)
                    cells[2].number_format = USD_FMT_NEG_RED

                out_ws.column_dimensions[get_column_letter(t_pkr)].width = 22
                out_ws.column_dimensions[get_column_letter(t_pct)].width = 14
                out_ws.column_dimensions[get_column_letter(t_usd)].width = 22
                company_boxes.append((t_pkr, t_usd))
                out_col += 3

            month_boxes.append((blk_start, out_col - 1))
            for c in range(blk_start, out_col):
                cell = out_ws.cell(row=1, column=c)
                cell.fill = _fill(pal_ban)
                cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            out_ws.cell(row=1, column=blk_start, value="Comparison in Amount — " + st["title"])
            out_ws.merge_cells(start_row=1, start_column=blk_start, end_row=1, end_column=out_col - 1)

    # ---- LAST column pair: Combined Total (PKR + USD) ---------------------
    if add_combined:
        c_pos, c_usd_pos = out_col, out_col + 1
        tot_cols += [c_pos, c_usd_pos]
        first = cols[0]
        _n, c_name_row = column_header_info(ws, first, header_row)
        if len(use_cols) > 1:
            c_head = "Combined Total — all selected companies" + (" & months" if len(files_bytes) > 1 else "")
        else:
            c_head = "Total — all selected months"
        out_ws.cell(row=c_name_row + RO, column=c_pos, value=c_head + " (PKR)")
        out_ws.cell(row=c_name_row + RO, column=c_usd_pos, value=c_head + " (USD)")
        style_header(c_pos, "tot")
        style_header(c_usd_pos, "tot")
        for _cp in (c_pos, c_usd_pos):
            _c = out_ws.cell(row=c_name_row + RO, column=_cp)
            _c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        _hr = out_ws.row_dimensions[c_name_row + RO]
        _hr.height = max(_hr.height or 0, 40)
        c_letter = get_column_letter(c_pos)
        c_num_len = 0
        for r in range(header_row + 1, max_row + 1):
            refs = []
            usd_refs = []
            for c in use_cols:
                for mi in range(len(ws_values_list)):
                    pos = positions.get((c, mi))
                    if pos and out_ws.cell(row=r + RO, column=pos).value is not None:
                        refs.append(f"{get_column_letter(pos)}{r + RO}")
                        usd_refs.append(f"{get_column_letter(pos + 1)}{r + RO}")
            kind = row_kind.get(r)
            key = kind == "key"
            font = Font(name="Arial", size=10 if key else 9,
                        bold=key or kind in ("heading", "total"), color="FF000000")
            for cp in (c_pos, c_usd_pos):
                cell = out_ws.cell(row=r + RO, column=cp)
                cell.fill = _fill(XL_COLORS["tot_key" if key else "tot_cell"])
                cell.font = font
                cell.alignment = Alignment(horizontal="right")
            if not refs:
                continue
            out_ws.cell(row=r + RO, column=c_pos, value="=" + "+".join(refs)).number_format = PKR_FMT_NEG_RED
            out_ws.cell(row=r + RO, column=c_usd_pos, value="=" + "+".join(usd_refs)).number_format = USD_FMT_NEG_RED
            tv = sum(sum_cell_across_files(ws_values_list, r, c)[0] for c in use_cols)
            c_num_len = max(c_num_len, len(f"{tv:,.2f}"), len(f"{tv / min(rates):,.2f}") + 1)
        out_ws.column_dimensions[c_letter].width = min(max(30, c_num_len + 5), 60)
        out_ws.column_dimensions[get_column_letter(c_usd_pos)].width = min(max(30, c_num_len + 5), 60)
        for c in (c_pos, c_usd_pos):
            cell = out_ws.cell(row=1, column=c)
            cell.fill = _fill(XL_COLORS["month_bg"])
            cell.font = Font(name="Arial", size=12, bold=True, color="FFFFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        out_ws.cell(row=1, column=c_pos, value="GRAND TOTAL")
        out_ws.merge_cells(start_row=1, start_column=c_pos, end_row=1, end_column=c_usd_pos)
        company_boxes.append((c_pos, c_usd_pos))
        month_boxes.append((c_pos, c_usd_pos))
        out_col += 2

    # ---- borders -----------------------------------------------------------
    for c in list(range(label_col_count + 1, out_col)):
        for r in range(1, last_out_row + 1):
            if r == 1:
                continue
            out_ws.cell(row=r, column=c).border = light_border

    dark = XL_COLORS["dark_line"]

    def box(c1, c2, r1, r2, style):
        s = Side(style=style, color=dark)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = out_ws.cell(row=r, column=c)
                b = cell.border
                cell.border = Border(
                    left=s if c == c1 else b.left,
                    right=s if c == c2 else b.right,
                    top=s if r == r1 else b.top,
                    bottom=s if r == r2 else b.bottom,
                )

    for c1, c2 in company_boxes:
        box(c1, c2, 2, last_out_row, "medium")
    for c1, c2 in month_boxes:
        box(c1, c2, 1, last_out_row, "thick")
    box(1, label_col_count, 1, last_out_row, "medium")
    for r, kind in row_kind.items():
        if kind == "key":
            s = Side(style="medium", color=dark)
            for c in range(1, out_col):
                cell = out_ws.cell(row=r + RO, column=c)
                b = cell.border
                cell.border = Border(left=b.left, right=b.right, top=s, bottom=s)

    # ---- notes -------------------------------------------------------------
    all_names = []
    for c in cols:
        nm, _nr = column_header_info(ws, c, header_row)
        all_names.append(nm or f"Column {c}")
    notes = ["Selected (" + str(len(all_names)) + "): " + ", ".join(all_names),
             "USD rate per sheet (PKR per 1 USD): " + ", ".join(f"{mn} = {rt:,.2f}" for mn, rt in zip(month_names, rates)),
             "Months (left to right): " + ", ".join(month_names)]
    if compare:
        notes.append("Comparison in Amount: every pair of selected months is compared. Difference = later month minus earlier month (PKR). "
                     "% = Difference ÷ |earlier month PKR|. USD difference = USD of the later month minus USD of the earlier month, "
                     "each converted at its own sheet's rate. The OVERALL block compares the last selected month with the first.")
        notes.append("TOTAL block (last 3 columns of each comparison): Difference PKR and Difference USD = sum of every selected "
                     "company's difference; % = total Difference PKR ÷ |sum of the earlier month PKR of the same companies|. "
                     "Companies already inside a selected group total / Grand Total are not added twice.")
    if add_combined:
        nm_of = lambda c: column_header_info(ws, c, header_row)[0] or f"Column {c}"
        line = "Grand Total adds: " + ", ".join(nm_of(c) for c in use_cols)
        if skipped_cols:
            line += "  |  not added again (already inside a selected group total): " + ", ".join(nm_of(c) for c in skipped_cols)
        notes.append(line)
    last_tbl_col = max(out_col - 1, label_col_count + 1)
    tot_w = 0.0
    for cc in range(1, last_tbl_col + 1):
        tot_w += out_ws.column_dimensions[get_column_letter(cc)].width or 8.43
    note_font = Font(name="Arial", size=10, bold=False, color="FF222222")
    head_font = Font(name="Arial", size=10, bold=True, color="FF000000")
    nr = last_out_row + 2
    out_ws.cell(row=nr, column=1, value="Notes").font = head_font
    nr += 1
    for note in notes:
        out_ws.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=last_tbl_col)
        c = out_ws.cell(row=nr, column=1, value=note)
        c.font = note_font
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        lines = max(1, math.ceil(len(note) * 1.1 / max(tot_w, 30)))
        out_ws.row_dimensions[nr].height = 14 * lines + 4
        nr += 1

    # ---- collapsible rows, freeze panes ----------------------------------
    out_ws.sheet_properties.outlinePr = Outline(summaryBelow=True, summaryRight=False)
    levels, closers = compute_row_outline(ws, header_row + 1, max_row)
    for r, lvl in levels.items():
        dim = out_ws.row_dimensions[r + RO]
        dim.outline_level = lvl
        if lvl >= OUTLINE_COLLAPSE_FROM_LEVEL:
            dim.hidden = True
    for r, child_lvl in closers.items():
        if child_lvl >= OUTLINE_COLLAPSE_FROM_LEVEL:
            out_ws.row_dimensions[r + RO].collapsed = True

    out_ws.freeze_panes = out_ws.cell(row=header_row + 1 + RO, column=label_col_count + 1).coordinate
    out_ws.sheet_view.zoomScale = 90

    out = io.BytesIO()
    out_wb.save(out)
    out.seek(0)
    return out.read()

