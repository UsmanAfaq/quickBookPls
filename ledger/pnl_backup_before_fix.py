
"""P&L hierarchy extraction (Revenue -> COGS -> Gross Profit -> ... -> Net) and USD conversion."""
from .sheet_reader import (PNL_TOTAL_PREFIXES, find_row_label, find_row_label_col,
                           find_total_column, map_column, norm_label,
                           sum_cell_across_files)

# QuickBooks P&L headings that are structural labels, not individual expenses.
PNL_HEADINGS = {
    "ordinary income/expense", "income", "revenue", "cost of goods sold - cogs",
    "cost of goods sold", "network subscription", "cost of services sold",
    "marketing expenses", "expense", "network subscriptions",
    "branding / advertising / promo", "finance cost", "operating expenses",
    "rent & maintenance", "staff salaries & other benefits", "taxation",
    "utility bills", "other income/expense", "other income", "other expense",
    "management expenses",
}

NUMERIC_KEYS = [
    "total_revenue", "total_cogs", "gross_profit", "total_operating_expense",
    "net_operational_income", "total_other_income", "total_other_expense",
    "net_other_income", "net_income", "total_expense",
]
ITEM_KEYS = [
    "revenue_items", "cogs_items", "operating_expense_items",
    "other_income_items", "other_expense_items",
]


def _is_cogs_marker(low):
    """COGS heading is spelled differently month to month."""
    return low.startswith("cost of goods sold") or low == "cogs"


def _numeric_sum(ws_values_list, row, cols):
    total = 0.0
    found = False
    for col in cols:
        value, ok = sum_cell_across_files(ws_values_list, row, col)
        if ok:
            total += float(value)
            found = True
    return total, found


def _pnl_rows(ws, ws_values_list, cols, header_row):
    """Read a QuickBooks P&L into the financial hierarchy (see module doc)."""
    if isinstance(cols, int):
        cols = [cols]

    state = None
    cogs_subsection = None
    operating_subsection = None
    other_subsection = None

    cogs_rows = []
    cogs_row_seen = {}
    cogs_detail_done = False

    sections = {"revenue": [], "cogs": [], "operating_expense": [],
                "other_income": [], "other_expense": []}

    for row in range(header_row + 1, ws.max_row + 1):
        label = find_row_label(ws, row)
        low = norm_label(label)

        if state == "cogs" and not cogs_detail_done and not _is_cogs_marker(low):
            # The COGS dropdown must be based on the sheet's row structure,
            # not on whether this particular company has a value in the row.
            # Otherwise zero/blank items disappear for some companies.
            if label != f"Row {row}":
                detail_total, detail_found = _numeric_sum(ws_values_list, row, cols)
                if low.startswith("total "):
                    detail_kind = "total"
                elif low in PNL_HEADINGS:
                    detail_kind = "heading"
                else:
                    detail_kind = "item"
                label_col = find_row_label_col(ws, row) or 5
                n = cogs_row_seen.get((detail_kind, low), 0)
                cogs_row_seen[(detail_kind, low)] = n + 1
                cogs_rows.append({
                    "key": f"{detail_kind}|{low}|{n}",
                    "kind": detail_kind,
                    "label": label,
                    "level": max(0, label_col - 5),
                    "rs": float(detail_total) if detail_found else 0.0,
                })
            if low == "total cost of services sold":
                cogs_detail_done = True

        # Revenue section heading
        if low == "revenue":
            state = "revenue"
            cogs_subsection = operating_subsection = other_subsection = None
            continue

        # QuickBooks may export account labels as REVENUE:Account Name.
        # Treat these as revenue accounts and remove the REVENUE: prefix.
        if low.startswith("revenue:") or low.startswith("revenue :"):
            state = "revenue"
            cogs_subsection = operating_subsection = other_subsection = None
            label = label.split(":", 1)[1].strip()
            low = norm_label(label)

        if _is_cogs_marker(low):
            state = "cogs"
            cogs_subsection = None
            continue

        if low == "expense":
            state = "operating_expense"
            operating_subsection = None
            continue

        if low == "other income/expense":
            state = "other"
            other_subsection = None
            continue

        if low == "other income":
            if state == "other":
                other_subsection = "other_income"
            continue

        if low == "other expense":
            if state == "other":
                other_subsection = "other_expense"
            continue

        if low == "total income":
            continue

        if low.startswith(PNL_TOTAL_PREFIXES):
            continue

        if low in PNL_HEADINGS:
            if state == "cogs":
                if low == "network subscription":
                    cogs_subsection = "Network Subscription"
                elif low == "cost of services sold":
                    cogs_subsection = "Cost of Services Sold"
                elif low == "marketing expenses":
                    cogs_subsection = "Marketing Expenses"
            elif state == "operating_expense":
                operating_subsection = label
            elif state == "other":
                if low == "management expenses":
                    other_subsection = "other_expense"
            continue

        row_total, found = _numeric_sum(ws_values_list, row, cols)
        if not found or row_total == 0:
            continue

        if state == "revenue":
            sections["revenue"].append({
                "category": "Revenue",
                "label": label,
                "rs": row_total
            })

        elif state == "cogs":
            sections["cogs"].append({
                "category": "COGS",
                "subsection": cogs_subsection or "COGS",
                "label": label,
                "rs": row_total
            })

        elif state == "operating_expense":
            sections["operating_expense"].append({
                "category": "Operating Expense",
                "subsection": operating_subsection or "Operating Expenses",
                "label": label,
                "rs": row_total
            })

        elif state == "other" and other_subsection == "other_income":
            sections["other_income"].append({
                "category": "Other Income",
                "label": label,
                "rs": row_total
            })

        elif state == "other" and other_subsection == "other_expense":
            sections["other_expense"].append({
                "category": "Other Expense",
                "subsection": other_subsection,
                "label": label,
                "rs": row_total
            })

    total_revenue = sum(x["rs"] for x in sections["revenue"])
    total_cogs = sum(x["rs"] for x in sections["cogs"])
    gross_profit = total_revenue - total_cogs
    total_operating_expense = sum(x["rs"] for x in sections["operating_expense"])
    net_operational_income = gross_profit - total_operating_expense
    total_other_income = sum(x["rs"] for x in sections["other_income"])
    total_other_expense = sum(x["rs"] for x in sections["other_expense"])
    net_other_income = total_other_income - total_other_expense
    net_income = net_operational_income + net_other_income

    return {
        "revenue_items": sections["revenue"],
        "expense_items": sections["cogs"] + sections["operating_expense"],
        "cogs_items": sections["cogs"],
        "cogs_rows": cogs_rows,
        "operating_expense_items": sections["operating_expense"],
        "other_income_items": sections["other_income"],
        "other_expense_items": sections["other_expense"],
        "total_revenue": total_revenue,
        "total_cogs": total_cogs,
        "gross_profit": gross_profit,
        "total_operating_expense": total_operating_expense,
        "net_operational_income": net_operational_income,
        "total_other_income": total_other_income,
        "total_other_expense": total_other_expense,
        "net_other_income": net_other_income,
        "net_income": net_income,
        "total_expense": total_cogs + total_operating_expense,
    }


def _month_breakdowns(ws, ws_values_list, cols, header_row):
    """One P&L per file. Each file is read with ITS OWN layout (rows, header row,
    TOTAL column); the selected columns are re-located in it by header text."""
    if isinstance(cols, int):
        cols = [cols]
    total_src = find_total_column(ws)[0]
    months = []
    for mv in ws_values_list:
        if mv is ws:
            m_cols, m_header = list(cols), header_row
        else:
            total_dst, m_header = find_total_column(mv)
            m_cols = [c for c in (map_column(ws, col, total_src, mv, total_dst) for col in cols)
                      if c is not None]
            m_header = m_header or header_row
        months.append(_pnl_rows(mv, [mv], m_cols, m_header))
    return months


def _merge_months(months):
    total = {}
    for key in ITEM_KEYS:
        total[key] = _merge_pnl_items(months, key)
    total["cogs_rows"] = _merge_cogs_rows(months)
    total["expense_items"] = total["cogs_items"] + total["operating_expense_items"]
    for key in NUMERIC_KEYS:
        total[key] = sum(float(m.get(key, 0) or 0) for m in months)
    return total


def company_breakdown(ws, ws_values, cols, header_row):
    ws_values_list = ws_values if isinstance(ws_values, list) else [ws_values]
    months = _month_breakdowns(ws, ws_values_list, cols, header_row)
    return months[0] if len(months) == 1 else _merge_months(months)


def _merge_pnl_items(months, key):
    merged = {}
    order = []
    for month in months:
        for item in month.get(key, []):
            label = item["label"]
            if label not in merged:
                merged[label] = dict(item)
                merged[label]["rs"] = 0.0
                order.append(label)
            merged[label]["rs"] += float(item["rs"])
    return [merged[label] for label in order]


def _merge_cogs_rows(months):
    merged = {}
    order = []
    for month in months:
        for row in month.get("cogs_rows", []):
            key = row["key"]
            if key not in merged:
                merged[key] = dict(row)
                merged[key]["rs"] = 0.0
                order.append(key)
            merged[key]["rs"] += float(row["rs"])
    return [merged[k] for k in order]


def company_breakdown_monthly(ws, ws_values_list, cols, header_row):
    """Selected entities kept separate, each with every month separate."""
    if isinstance(cols, int):
        cols = [cols]
    results = []
    for col in cols:
        months = _month_breakdowns(ws, ws_values_list, [col], header_row)
        results.append({"col": col, "months": months, "total": _merge_months(months)})
    return results


# ---- USD conversion ------------------------------------------------------

def add_usd_to_pnl(pnl, rate):
    for key in ITEM_KEYS + ["cogs_rows"]:
        for item in pnl.get(key, []):
            item["usd"] = item["rs"] / rate
    for item in pnl.get("expense_items", []):
        item["usd"] = item["rs"] / rate
    for key in NUMERIC_KEYS:
        pnl[f"{key}_usd"] = pnl.get(key, 0) / rate
    return pnl


def add_usd_to_selection(selection, rates):
    """USD per month at THAT month's own rate; the all-months total is the
    sum of monthly USD figures (never total Rs / one rate)."""
    months = selection["months"]
    for m, r in zip(months, rates):
        m["rate"] = r
        add_usd_to_pnl(m, r)
    tot = selection["total"]
    for key in ITEM_KEYS:
        by_label = {}
        for m in months:
            for it in m.get(key, []):
                by_label[it["label"]] = by_label.get(it["label"], 0.0) + it.get("usd", 0.0)
        for it in tot.get(key, []):
            it["usd"] = by_label.get(it["label"], 0.0)
    by_key = {}
    for m in months:
        for row in m.get("cogs_rows", []):
            by_key[row["key"]] = by_key.get(row["key"], 0.0) + row.get("usd", 0.0)
    for row in tot.get("cogs_rows", []):
        row["usd"] = by_key.get(row["key"], 0.0)
    for k in NUMERIC_KEYS:
        tot[f"{k}_usd"] = sum(m.get(f"{k}_usd", 0.0) for m in months)
    rs_sum = sum(m.get("total_revenue", 0.0) for m in months)
    usd_sum = tot["total_revenue_usd"]
    tot["eff_rate"] = (rs_sum / usd_sum) if usd_sum else (sum(rates) / len(rates))
    return selection
