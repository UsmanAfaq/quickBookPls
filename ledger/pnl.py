"""P&L hierarchy extraction (Revenue -> COGS -> Gross Profit -> ... -> Net) and USD conversion.

QuickBooks flat exports commonly store account rows as CATEGORY:ACCOUNT NAME.
Only account rows are added. Heading and subtotal rows are never added again.
"""
from .sheet_reader import (
    PNL_TOTAL_PREFIXES,
    find_row_label,
    find_row_label_col,
    find_total_column,
    map_column,
    norm_label,
    sum_cell_across_files,
)

# Structural labels used by the legacy indented/outline export format.
PNL_HEADINGS = {
    "ordinary income/expense", "income", "revenue",
    "cost of goods sold - cogs", "cost of goods sold", "network subscription",
    "cost of services sold", "marketing expenses", "expense",
    "network subscriptions", "branding / advertising / promo", "finance cost",
    "operating expenses", "rent & maintenance", "rent & pitb utilities",
    "staff salaries & other benefits", "taxation", "utility bills",
    "other income/expense", "other income", "other expense", "other expenses",
    "management expenses",
}

# QuickBooks account-prefix -> dashboard section.
_REVENUE_PREFIXES = {"revenue", "income"}
_COGS_PREFIXES = {
    "cost of goods sold", "cost of goods sold - cogs", "cogs",
    "marketing expenses", "network subscription", "network subscriptions",
    "cost of services sold",
}
_OPERATING_PREFIXES = {
    "branding / advertising / promo", "finance cost", "operating expenses",
    "rent & pitb utilities", "rent & maintenance",
    "staff salaries & other benefits", "taxation", "utility bills",
    "management expenses",
}
_OTHER_INCOME_PREFIXES = {"other income"}
_OTHER_EXPENSE_PREFIXES = {"other expense", "other expenses"}

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


def _account_parts(label):
    """Return (normalized category prefix, account path) for CATEGORY:ACCOUNT rows."""
    if ":" not in label:
        return None, None
    prefix, account = label.split(":", 1)
    prefix = norm_label(prefix).rstrip(": ")
    account = account.strip()
    if not prefix or not account:
        return None, None
    return prefix, account


def _append_item(target, category, label, amount, subsection=None):
    item = {"category": category, "label": label, "rs": float(amount)}
    if subsection:
        item["subsection"] = subsection
    target.append(item)


def _pnl_rows(ws, ws_values_list, cols, header_row):
    """Read P&L account rows without adding QuickBooks subtotal rows twice."""
    if isinstance(cols, int):
        cols = [cols]

    sections = {
        "revenue": [], "cogs": [], "operating_expense": [],
        "other_income": [], "other_expense": [],
    }
    cogs_rows = []
    cogs_seen = {}

    for row in range((header_row or 1) + 1, ws.max_row + 1):
        raw_label = find_row_label(ws, row)
        label = raw_label.strip()
        low = norm_label(label)

        # Flat QuickBooks export: account detail rows have a category prefix.
        prefix, account = _account_parts(label)
        if prefix is not None:
            amount, found = _numeric_sum(ws_values_list, row, cols)
            if not found:
                continue
            # Preserve zero-value account rows in the COGS detail structure,
            # but do not create noisy zero rows in dashboard expense lists.
            if prefix in _REVENUE_PREFIXES:
                if amount:
                    _append_item(sections["revenue"], "Revenue", account, amount)
            elif prefix in _COGS_PREFIXES:
                subsection = prefix.title()
                if amount:
                    _append_item(sections["cogs"], "COGS", account, amount, subsection)
                key_count = cogs_seen.get((prefix, account), 0)
                cogs_seen[(prefix, account)] = key_count + 1
                cogs_rows.append({
                    "key": f"item|{prefix}|{norm_label(account)}|{key_count}",
                    "kind": "item", "label": account,
                    "level": max(0, label.count(":") - 1),
                    "rs": float(amount),
                })
            elif prefix in _OPERATING_PREFIXES:
                if amount:
                    _append_item(
                        sections["operating_expense"], "Operating Expense",
                        account, amount, prefix.title(),
                    )
            elif prefix in _OTHER_INCOME_PREFIXES:
                if amount:
                    _append_item(sections["other_income"], "Other Income", account, amount)
            elif prefix in _OTHER_EXPENSE_PREFIXES:
                if amount:
                    _append_item(
                        sections["other_expense"], "Other Expense", account,
                        amount, "Other Expenses",
                    )
            # Unknown prefixed accounts are intentionally not guessed into a
            # financial category. This prevents accidental COGS/expense inflation.
            continue

        # Legacy indented export fallback. These rows are headings/totals only;
        # the flat export's unprefixed rows are not account detail rows.
        # No subtotal row is added to a financial section.
        if low.startswith(PNL_TOTAL_PREFIXES) or low in PNL_HEADINGS:
            continue

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
        "total_expense": total_cogs + total_operating_expense + total_other_expense,
    }


def _month_breakdowns(ws, ws_values_list, cols, header_row):
    """Read each monthly workbook using its own layout and matching columns."""
    if isinstance(cols, int):
        cols = [cols]
    total_src = find_total_column(ws)[0]
    months = []
    for mv in ws_values_list:
        if mv is ws:
            m_cols, m_header = list(cols), header_row
        else:
            total_dst, m_header = find_total_column(mv)
            m_cols = [
                c for c in (map_column(ws, col, total_src, mv, total_dst) for col in cols)
                if c is not None
            ]
            m_header = m_header or header_row
        months.append(_pnl_rows(mv, [mv], m_cols, m_header))
    return months


def _merge_pnl_items(months, key):
    merged = {}
    order = []
    for month in months:
        for item in month.get(key, []):
            # Include subsection in identity so same account text in two sections
            # does not merge incorrectly.
            identity = (item.get("subsection", ""), item["label"])
            if identity not in merged:
                merged[identity] = dict(item)
                merged[identity]["rs"] = 0.0
                order.append(identity)
            merged[identity]["rs"] += float(item["rs"])
    return [merged[k] for k in order]


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


def _merge_months(months):
    total = {key: _merge_pnl_items(months, key) for key in ITEM_KEYS}
    total["cogs_rows"] = _merge_cogs_rows(months)
    total["expense_items"] = total["cogs_items"] + total["operating_expense_items"]
    for key in NUMERIC_KEYS:
        total[key] = sum(float(m.get(key, 0) or 0) for m in months)
    return total


def company_breakdown(ws, ws_values, cols, header_row):
    ws_values_list = ws_values if isinstance(ws_values, list) else [ws_values]
    months = _month_breakdowns(ws, ws_values_list, cols, header_row)
    return months[0] if len(months) == 1 else _merge_months(months)


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
    if not rate:
        raise ValueError("USD conversion rate must be a non-zero number")
    for key in ITEM_KEYS + ["cogs_rows"]:
        for item in pnl.get(key, []):
            item["usd"] = item["rs"] / rate
    for item in pnl.get("expense_items", []):
        item["usd"] = item["rs"] / rate
    for key in NUMERIC_KEYS:
        pnl[f"{key}_usd"] = pnl.get(key, 0) / rate
    return pnl


def add_usd_to_selection(selection, rates):
    """Convert each month at its own rate; total USD is sum of monthly USD."""
    months = selection["months"]
    if len(months) != len(rates):
        raise ValueError("The number of exchange rates must match the number of months")
    for month, rate in zip(months, rates):
        month["rate"] = rate
        add_usd_to_pnl(month, rate)

    total = selection["total"]
    for key in ITEM_KEYS:
        by_identity = {}
        for month in months:
            for item in month.get(key, []):
                identity = (item.get("subsection", ""), item["label"])
                by_identity[identity] = by_identity.get(identity, 0.0) + item.get("usd", 0.0)
        for item in total.get(key, []):
            identity = (item.get("subsection", ""), item["label"])
            item["usd"] = by_identity.get(identity, 0.0)

    by_key = {}
    for month in months:
        for row in month.get("cogs_rows", []):
            by_key[row["key"]] = by_key.get(row["key"], 0.0) + row.get("usd", 0.0)
    for row in total.get("cogs_rows", []):
        row["usd"] = by_key.get(row["key"], 0.0)

    for key in NUMERIC_KEYS:
        total[f"{key}_usd"] = sum(month.get(f"{key}_usd", 0.0) for month in months)
    rs_sum = sum(month.get("total_revenue", 0.0) for month in months)
    usd_sum = total.get("total_revenue_usd", 0.0)
    total["eff_rate"] = rs_sum / usd_sum if usd_sum else (sum(rates) / len(rates) if rates else 0.0)
    return selection
