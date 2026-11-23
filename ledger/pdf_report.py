"""PDF reports: matplotlib charts + reportlab tables."""
import io

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (Image as RLImage, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

from .pnl import add_usd_to_selection, company_breakdown_monthly
from .rates import as_rates, rates_text, same_rate
from .sheet_reader import (PNL_TOTAL_PREFIXES, PROMINENT_LABELS, build_reference,
                           find_row_label_col, load_sheet, load_values_sheet, norm_label)

PDF_KEY_LABELS = PROMINENT_LABELS


def _build_report_data(contents, rate, col_list, selection_names, month_names=None):
    ws, total_col, header_row, _last = build_reference(contents)
    ws_values_list = [load_values_sheet(c) for c in contents]
    month_names = month_names or [f"Month {i + 1}" for i in range(len(contents))]
    selections = company_breakdown_monthly(ws, ws_values_list, col_list, header_row)
    for idx, selection in enumerate(selections):
        selection["name"] = (selection_names[idx] if idx < len(selection_names)
                             else f"Column {selection['col']}")
        for midx, month in enumerate(selection["months"]):
            month["name"] = month_names[midx]
        add_usd_to_selection(selection, as_rates(rate, len(contents)))
    return selections, month_names


def _usd_fmt(v, _p):
    return f"-${abs(v):,.0f}" if v < 0 else f"${v:,.0f}"


def _pkr_fmt(v, _p):
    return f"-Rs {abs(v):,.0f}" if v < 0 else f"Rs {v:,.0f}"


def _fig_png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _chart_png_monthly(selection, rates):
    rate = same_rate(rates)
    names = [m["name"] for m in selection["months"]]
    series = [
        ("Revenue", "total_revenue_usd"), ("COGS", "total_cogs_usd"),
        ("Gross Profit", "gross_profit_usd"), ("Operating Expense", "total_operating_expense_usd"),
        ("Net Operational Income", "net_operational_income_usd"), ("Net Profit / Loss", "net_income_usd"),
    ]
    fig, ax = plt.subplots(figsize=(10, 4.5), dpi=150)
    x = list(range(len(names)))
    for label, key in series:
        ax.plot(x, [m[key] for m in selection["months"]], marker="o", label=label)
    ax.set_xticks(x, names)
    ax.set_ylabel("USD ($)")
    ax.set_title(f"{selection['name']} — Monthly P&L   " + (f"(left axis: USD $  |  right axis: PKR Rs @ {rate:,.2f} per USD)" if rate else "(USD $ — each sheet at its own rate)"), fontsize=9)
    ax.grid(axis="y", alpha=0.25)
    ax.axhline(0, color="#444", linewidth=0.8)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_usd_fmt))
    if rate:
        sec = ax.secondary_yaxis("right", functions=(lambda v: v * rate, lambda v: v / rate))
        sec.set_ylabel("PKR (Rs)")
        sec.yaxis.set_major_formatter(plt.FuncFormatter(_pkr_fmt))
    ax.legend(fontsize=8, ncol=2)
    fig.tight_layout()
    return _fig_png(fig)


def _chart_png_revenue_items(selection, rates):
    """Vertical bars: one group per revenue item, one bar per month."""
    rate = same_rate(rates)
    months = selection["months"]
    labels, seen = [], set()
    for m in months:
        for item in m["revenue_items"]:
            if item["label"] not in seen:
                seen.add(item["label"])
                labels.append(item["label"])
    if not labels:
        return None
    totals = {l: sum({i["label"]: i["usd"] for i in m["revenue_items"]}.get(l, 0) for m in months) for l in labels}
    labels = sorted(labels, key=lambda l: -abs(totals[l]))[:12]
    n_m = max(len(months), 1)
    fig_w = max(9, min(16, 3 + len(labels) * (0.55 + 0.28 * n_m)))
    fig, ax = plt.subplots(figsize=(fig_w, 5.4), dpi=150)
    x = list(range(len(labels)))
    width = 0.8 / n_m
    for idx, m in enumerate(months):
        values = {i["label"]: i["usd"] for i in m["revenue_items"]}
        vals = [values.get(l, 0) for l in labels]
        offsets = [v + (idx - (n_m - 1) / 2) * width for v in x]
        bars = ax.bar(offsets, vals, width=width * 0.92, label=m["name"])
        ax.bar_label(bars, labels=[(f"-${abs(v):,.0f}" if v < 0 else f"${v:,.0f}") if v else "" for v in vals],
                     rotation=0, padding=2, fontsize=6.5)
    short = [l if len(l) <= 30 else l[:29] + "…" for l in labels]
    ax.set_xticks(x, short, rotation=35, ha="right", fontsize=8)
    ax.set_ylabel("USD ($)")
    ax.margins(y=0.18)
    ax.set_title(f"{selection['name']} — Revenue by Ad / Revenue Item   " + (f"(left: USD $ | right: PKR Rs @ {rate:,.2f} per USD)" if rate else "(USD $ — each sheet at its own rate)"), fontsize=9)
    ax.grid(axis="y", alpha=0.25)
    ax.axhline(0, color="#444", linewidth=0.8)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(_usd_fmt))
    if rate:
        sec = ax.secondary_yaxis("right", functions=(lambda v: v * rate, lambda v: v / rate))
        sec.set_ylabel("PKR (Rs)")
        sec.yaxis.set_major_formatter(plt.FuncFormatter(_pkr_fmt))
    if len(months) > 1:
        ax.legend(fontsize=8)
    fig.tight_layout()
    return _fig_png(fig)


def _png_to_flowable(png_bytes, max_w, max_h):
    from reportlab.lib.utils import ImageReader
    iw, ih = ImageReader(io.BytesIO(png_bytes)).getSize()
    scale = min(max_w / iw, max_h / ih)
    return RLImage(io.BytesIO(png_bytes), width=iw * scale, height=ih * scale)


def _esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _new_doc(out):
    return SimpleDocTemplate(out, pagesize=landscape(A4), rightMargin=28, leftMargin=28,
                             topMargin=28, bottomMargin=28)


# ---------------------------------------------------------------------------
# Graphs / revenue half
# ---------------------------------------------------------------------------

def _report_story(contents, rate, col_list, selection_names, month_names=None, with_title=True):
    selections, month_names = _build_report_data(contents, rate, col_list, selection_names, month_names=month_names)
    rates = as_rates(rate, len(contents))
    rtxt = rates_text(month_names, rates)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("ReportTitle", parent=styles["Title"], fontSize=20, spaceAfter=8)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=14, spaceBefore=10, spaceAfter=6)
    small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=8)
    story = []
    if with_title:
        story += [
            Paragraph("Rs → USD P&amp;L Report", title),
            Paragraph("Months: " + ", ".join(month_names), small),
            Paragraph("Exchange rate: " + rtxt, small),
            Spacer(1, 10),
        ]
    else:
        story += [Paragraph("Graphs &amp; Revenue Sources", title),
                  Paragraph("Charts show USD ($) — " + rtxt, small),
                  Spacer(1, 6)]
    max_w = landscape(A4)[0] - 56

    for idx, selection in enumerate(selections):
        story.append(Paragraph(_esc(f"{idx + 1}. {selection['name']}"), h2))
        story.append(_png_to_flowable(_chart_png_monthly(selection, rates), max_w, 3.7 * inch))
        item_png = _chart_png_revenue_items(selection, rates)
        if item_png:
            story.append(Spacer(1, 8))
            story.append(_png_to_flowable(item_png, max_w, 4.3 * inch))

        labels, seen = [], set()
        for m in selection["months"]:
            for item in m["revenue_items"]:
                if item["label"] not in seen:
                    seen.add(item["label"])
                    labels.append(item["label"])
        if labels:
            multi = len(selection["months"]) > 1
            head = ["Ad / Revenue Item"] + [m["name"] + " (USD)" for m in selection["months"]]
            head += (["All Months (USD)", "All Months (PKR)"] if multi else ["PKR"])
            data = [head]
            total_map = {i["label"]: i for i in selection["total"]["revenue_items"]}
            for label in labels:
                row = [Paragraph(_esc(label), small)]
                for m in selection["months"]:
                    mm = {i["label"]: i["usd"] for i in m["revenue_items"]}
                    row.append(f"${mm.get(label, 0):,.2f}")
                t_item = total_map.get(label, {})
                if multi:
                    row.append(f"${t_item.get('usd', 0):,.2f}")
                row.append(f"Rs {t_item.get('rs', 0):,.2f}")
                data.append(row)
            t = Table(data, repeatRows=1)
            cmds = [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#efe9d8")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cdc4a8")),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("BACKGROUND", (-1, 1), (-1, -1), colors.HexColor("#eef7f1")),
            ]
            for ri, rowv in enumerate(data[1:], start=1):
                for ci, v in enumerate(rowv[1:], start=1):
                    if isinstance(v, str) and ("-$" in v or "Rs -" in v):
                        cmds.append(("TEXTCOLOR", (ci, ri), (ci, ri), colors.HexColor("#d00000")))
            t.setStyle(TableStyle(cmds))
            story.extend([Spacer(1, 8), t])
        story.append(PageBreak())
    if story and isinstance(story[-1], PageBreak):
        story.pop()
    return story


# ---------------------------------------------------------------------------
# P&L statement half
# ---------------------------------------------------------------------------

def _selection_story(files_bytes, rate, cols, names, month_names, with_title=True):
    ws, _, total_col, header_row = load_sheet(files_bytes[0])
    wsv_list = [load_values_sheet(fb) for fb in files_bytes]
    if not month_names or len(month_names) != len(files_bytes):
        month_names = [f"File {i + 1}" for i in range(len(files_bytes))]
    multi = len(wsv_list) > 1
    rates = as_rates(rate, len(files_bytes))

    styles = getSampleStyleSheet()
    title_st = ParagraphStyle("PT", parent=styles["Title"], fontSize=16, spaceAfter=4)
    small = ParagraphStyle("PS", parent=styles["BodyText"], fontSize=8)
    story = []
    if with_title:
        story += [
            Paragraph("Rs → USD P&amp;L Statement", title_st),
            Paragraph("Months: " + ", ".join(month_names) +
                      " &nbsp;|&nbsp; Exchange rate: " + rates_text(month_names, rates), small),
            Spacer(1, 8),
        ]

    C_PKR_H, C_USD_H, C_TOT_H = colors.HexColor("#1F6F43"), colors.HexColor("#1F4E8C"), colors.HexColor("#8A5A00")
    C_PKR, C_USD, C_TOT = colors.HexColor("#EEF7F1"), colors.HexColor("#EEF3FB"), colors.HexColor("#FFF8E1")
    C_PKR_K, C_USD_K, C_TOT_K = colors.HexColor("#BFE3CB"), colors.HexColor("#C4D8F3"), colors.HexColor("#FFE08A")
    C_RED, C_LINE = colors.HexColor("#D00000"), colors.HexColor("#D5D5D5")
    DARK = colors.HexColor("#2B2B2B")

    max_row = ws.max_row
    chunk = 3
    for si, col in enumerate(cols):
        sel_name = names[si] if si < len(names) else f"Column {col}"
        for start in range(0, len(wsv_list), chunk):
            idxs = list(range(start, min(start + chunk, len(wsv_list))))
            heading = f"{si + 1}. {sel_name}"
            if len(wsv_list) > chunk:
                heading += " — " + ", ".join(month_names[i] for i in idxs)
            story.append(Paragraph(_esc(heading), styles["Heading2"]))

            n_blocks = len(idxs) + (1 if multi else 0)
            head1 = ["Item"]
            head2 = [""]
            for i in idxs:
                head1 += [month_names[i], ""]
                head2 += ["PKR (Rs)", "USD ($)"]
            if multi:
                head1 += ["ALL SELECTED MONTHS", ""]
                head2 += ["PKR (Rs)", "USD ($)"]
            data = [head1, head2]
            last_col = 2 * n_blocks
            cmds = [
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (1, 0), (-1, 1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.25, C_LINE),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("BACKGROUND", (0, 0), (0, 1), colors.HexColor("#1B2A41")),
                ("TEXTCOLOR", (0, 0), (0, 1), colors.white),
                ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
                ("TEXTCOLOR", (1, 0), (-1, 1), colors.white),
            ]
            for b in range(n_blocks):
                c1, c2 = 1 + 2 * b, 2 + 2 * b
                is_tot = multi and b == n_blocks - 1
                cmds += [
                    ("SPAN", (c1, 0), (c2, 0)),
                    ("BACKGROUND", (c1, 0), (c2, 0), colors.HexColor("#1B2A41") if not is_tot else C_TOT_H),
                    ("BACKGROUND", (c1, 1), (c1, 1), C_TOT_H if is_tot else C_PKR_H),
                    ("BACKGROUND", (c2, 1), (c2, 1), C_TOT_H if is_tot else C_USD_H),
                    ("BOX", (c1, 0), (c2, -1), 1.2, DARK),
                ]
            cmds.append(("BOX", (0, 0), (0, -1), 1.2, DARK))

            for r in range(header_row + 1, max_row + 1):
                lc = find_row_label_col(ws, r)
                if not lc:
                    continue
                label = ws.cell(row=r, column=lc).value.strip()
                low = norm_label(label)
                vals = []
                for i in idxs:
                    v = wsv_list[i].cell(row=r, column=col).value if r <= wsv_list[i].max_row else None
                    vals.append(float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None)
                all_vals = []
                all_usd = []
                for wi, wv in enumerate(wsv_list):
                    v = wv.cell(row=r, column=col).value if r <= wv.max_row else None
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        all_vals.append(float(v))
                        all_usd.append(float(v) / rates[wi])
                is_heading = all(v is None for v in vals) and not all_vals
                is_key = low in PDF_KEY_LABELS
                is_total = low.startswith(PNL_TOTAL_PREFIXES)
                bold = is_heading or is_total or is_key
                st = ParagraphStyle(
                    f"L{lc}{int(bold)}{int(is_key)}", parent=small, fontSize=8.5 if is_key else 7,
                    leftIndent=max(0, lc - 2) * 9,
                    fontName="Helvetica-Bold" if bold else "Helvetica",
                )
                row = [Paragraph(_esc(label), st)]
                nums = []
                for vi, v in enumerate(vals):
                    if v is None:
                        row += ["", ""]
                        nums += [None, None]
                    else:
                        u = v / rates[idxs[vi]]
                        row += [f"{v:,.2f}", f"{u:,.2f}"]
                        nums += [v, u]
                if multi:
                    if all_vals:
                        tv = sum(all_vals)
                        tu = sum(all_usd)
                        row += [f"{tv:,.2f}", f"{tu:,.2f}"]
                        nums += [tv, tu]
                    else:
                        row += ["", ""]
                        nums += [None, None]
                data.append(row)
                ri = len(data) - 1
                for b in range(n_blocks):
                    c1, c2 = 1 + 2 * b, 2 + 2 * b
                    is_tot = multi and b == n_blocks - 1
                    if is_key:
                        cmds.append(("BACKGROUND", (c1, ri), (c1, ri), C_TOT_K if is_tot else C_PKR_K))
                        cmds.append(("BACKGROUND", (c2, ri), (c2, ri), C_TOT_K if is_tot else C_USD_K))
                    else:
                        cmds.append(("BACKGROUND", (c1, ri), (c1, ri), C_TOT if is_tot else C_PKR))
                        cmds.append(("BACKGROUND", (c2, ri), (c2, ri), C_TOT if is_tot else C_USD))
                if is_heading:
                    cmds.append(("BACKGROUND", (0, ri), (0, ri), colors.HexColor("#F1EEE2")))
                if bold:
                    cmds.append(("FONTNAME", (1, ri), (-1, ri), "Helvetica-Bold"))
                if is_key:
                    cmds.append(("FONTSIZE", (1, ri), (-1, ri), 8.5))
                    cmds.append(("BACKGROUND", (0, ri), (0, ri), colors.HexColor("#FFE9A8")))
                    cmds.append(("LINEABOVE", (0, ri), (-1, ri), 1.2, DARK))
                    cmds.append(("LINEBELOW", (0, ri), (-1, ri), 1.2, DARK))
                elif is_total:
                    cmds.append(("LINEABOVE", (0, ri), (-1, ri), 0.6, colors.HexColor("#96762f")))
                for ci, nv in enumerate(nums, start=1):
                    if nv is not None and nv < 0:
                        cmds.append(("TEXTCOLOR", (ci, ri), (ci, ri), C_RED))

            label_w = 190 if multi else 230
            num_w = (landscape(A4)[0] - 56 - label_w) / max(1, last_col)
            t = Table(data, colWidths=[label_w] + [num_w] * last_col, repeatRows=2)
            t.setStyle(TableStyle(cmds))
            story.append(t)
            story.append(PageBreak())

    if story and isinstance(story[-1], PageBreak):
        story.pop()
    return story


# ---------------------------------------------------------------------------
# Public builders
# ---------------------------------------------------------------------------

def build_pdf_report(contents, rate, col_list, selection_names, month_names=None):
    out = io.BytesIO()
    _new_doc(out).build(_report_story(contents, rate, col_list, selection_names, month_names=month_names))
    out.seek(0)
    return out.read()


def build_selection_pdf(files_bytes, rate, cols, names, month_names):
    out = io.BytesIO()
    _new_doc(out).build(_selection_story(files_bytes, rate, cols, names, month_names))
    out.seek(0)
    return out.read()


def build_full_pdf(files_bytes, rate, cols, names, month_names):
    """ONE file: P&L statement tables first, then graphs/revenue pages."""
    out = io.BytesIO()
    story = _selection_story(files_bytes, rate, cols, names, month_names)
    story.append(PageBreak())
    graph_months = month_names if month_names and len(month_names) == len(files_bytes) else None
    story += _report_story(files_bytes, rate, cols, names, month_names=graph_months, with_title=False)
    _new_doc(out).build(story)
    out.seek(0)
    return out.read()