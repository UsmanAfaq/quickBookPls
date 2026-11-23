"""JSON data routes: structure, breakdown, graph data, comparison."""
import json
from typing import List, Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from .pnl import add_usd_to_selection, company_breakdown, company_breakdown_monthly
from .rates import month_pairs, parse_rates
from .sheet_reader import (build_reference, find_companies, find_groups_and_standalone,
                           load_sheet, load_values_sheet)
from .usd_column import add_usd_column

router = APIRouter()

CMP_METRICS = [
    ("Revenue", "total_revenue"),
    ("COGS", "total_cogs"),
    ("Gross Profit", "gross_profit"),
    ("Operating Expense", "total_operating_expense"),
    ("Net Operational Income", "net_operational_income"),
    ("Other Income", "total_other_income"),
    ("Other Expense", "total_other_expense"),
    ("Net Other Income", "net_other_income"),
    ("Net Profit / Loss", "net_income"),
]


def _cmp_row(label, rs_list, rates, month_names):
    """Rs + USD per month, plus the difference for EVERY pair of months."""
    usd_list = [rs / r for rs, r in zip(rs_list, rates)]
    steps = []
    for a, b, title, ov in month_pairs(month_names):
        diff = rs_list[b] - rs_list[a]
        steps.append({
            "title": title,
            "overall": ov,
            "diff_rs": diff,
            "pct": (diff / abs(rs_list[a]) * 100) if rs_list[a] else None,
            "diff_usd": usd_list[b] - usd_list[a],
        })
    return {"label": label, "rs": rs_list, "usd": usd_list, "steps": steps}


@router.post("/compare_amount")
async def compare_amount_route(
    files: List[UploadFile] = File(...),
    cols: str = Form(...),
    names: str = Form(""),
    rates: str = Form(""),
    rate: Optional[float] = Form(None),
):
    contents = [await f.read() for f in files]
    if len(contents) < 2:
        raise HTTPException(status_code=400, detail="Select at least 2 months to compare.")
    try:
        rate_list = parse_rates(rates, len(contents), rate)
        ws, total_col, header_row, last_col = build_reference(contents)
        col_list = [int(c) for c in cols.split(",") if c.strip()]
        if not col_list or any(c < 1 or c > last_col for c in col_list):
            raise ValueError("Invalid company selection.")
        names_list = json.loads(names) if names else []
        if not isinstance(names_list, list) or len(names_list) != len(col_list):
            names_list = [f"Column {c}" for c in col_list]
        ws_values_list = [load_values_sheet(c) for c in contents]
    except (ValueError, TypeError, json.JSONDecodeError) as e:
        raise HTTPException(status_code=400, detail=str(e))

    month_names = [f.filename.rsplit(".", 1)[0] if f.filename else f"Month {i + 1}" for i, f in enumerate(files)]
    selections = company_breakdown_monthly(ws, ws_values_list, col_list, header_row)
    out = []
    for si, sel in enumerate(selections):
        summary = [
            _cmp_row(label, [float(m.get(key, 0) or 0) for m in sel["months"]], rate_list, month_names)
            for label, key in CMP_METRICS
        ]
        labels, seen = [], set()
        for m in sel["months"]:
            for it in m.get("revenue_items", []):
                if it["label"] not in seen:
                    seen.add(it["label"])
                    labels.append(it["label"])
        sources = []
        for label in labels:
            vals = []
            for m in sel["months"]:
                d = {i["label"]: float(i["rs"]) for i in m.get("revenue_items", [])}
                vals.append(d.get(label, 0.0))
            sources.append(_cmp_row(label, vals, rate_list, month_names))
        out.append({"name": names_list[si], "summary": summary, "sources": sources})
    return {"months": month_names, "rates": rate_list, "selections": out}


@router.post("/preview")
async def preview(file: UploadFile = File(...), rate: float = Form(...)):
    content = await file.read()
    try:
        _, summary = add_usd_column(content, rate)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"summary": summary}


@router.post("/companies")
async def companies(file: UploadFile = File(...)):
    content = await file.read()
    try:
        ws, _, total_col, header_row = load_sheet(content)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"companies": find_companies(ws, total_col, header_row)}


@router.post("/structure")
async def structure(files: List[UploadFile] = File(...)):
    # Teams come and go between months, so the list is built from EVERY selected file.
    contents = [await f.read() for f in files]
    try:
        ws, total_col, header_row, last_col = build_reference(contents)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    groups, standalone = find_groups_and_standalone(ws, total_col, last_col)
    return {"groups": list(groups.values()), "standalone": standalone, "grand_total_col": total_col}


@router.post("/company_breakdown")
async def company_breakdown_route(
    files: List[UploadFile] = File(...),
    cols: str = Form(...),
    rates: str = Form(""),
    rate: Optional[float] = Form(None),
    names: str = Form(""),
):
    contents = [await f.read() for f in files]
    try:
        ws, total_col, header_row, last_col = build_reference(contents)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    try:
        col_list = [int(c) for c in cols.split(",") if c.strip() != ""]
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid column selection.")
    if not col_list:
        raise HTTPException(status_code=400, detail="No columns selected.")
    for c in col_list:
        if c < 1 or c > last_col:
            raise HTTPException(status_code=400, detail="Invalid company column.")

    try:
        selection_names = json.loads(names) if names else []
    except json.JSONDecodeError:
        selection_names = []
    if not isinstance(selection_names, list) or len(selection_names) != len(col_list):
        selection_names = [f"Column {c}" for c in col_list]

    month_names = [f.filename or f"Month {i + 1}" for i, f in enumerate(files)]

    try:
        ws_values_list = [load_values_sheet(c) for c in contents]
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    result = company_breakdown(ws, ws_values_list, col_list, header_row)
    selections = company_breakdown_monthly(ws, ws_values_list, col_list, header_row)

    try:
        rate_list = parse_rates(rates, len(contents), rate)
    except ValueError:
        rate_list = None      # some sheet has no rate yet -> Rs only

    for idx, selection in enumerate(selections):
        selection["name"] = selection_names[idx]
        selection["months"] = [
            {"name": month_names[midx], **month}
            for midx, month in enumerate(selection["months"])
        ]
        if rate_list:
            add_usd_to_selection(selection, rate_list)

    result["selections"] = selections
    result["rates"] = rate_list
    result["month_names"] = month_names
    return result


@router.post("/graph_data")
async def graph_data_route(
    files: List[UploadFile] = File(...),
    rates: str = Form(""),
    rate: Optional[float] = Form(None),
    cols: str = Form(...),
    names: str = Form(""),
    requested_ad_names: str = Form(""),
    source: str = Form("revenue"),      # "revenue" -> ad/revenue sources, "cogs" -> COGS expense items,
                                        # "opex" -> Operating Expense + Other Expense items combined
):
    contents = [await f.read() for f in files]
    try:
        rate_list = parse_rates(rates, len(contents), rate)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    if not contents:
        raise HTTPException(status_code=400, detail="Invalid files.")

    try:
        ws, total_col, header_row, last_col = build_reference(contents)
        col_list = [int(c) for c in cols.split(",") if c.strip()]
        if not col_list or any(c < 1 or c > last_col for c in col_list):
            raise ValueError("Invalid company selection.")
        selection_names = json.loads(names) if names else []
        if not isinstance(selection_names, list) or len(selection_names) != len(col_list):
            selection_names = [f"Column {c}" for c in col_list]
        month_names = [f.filename.rsplit(".", 1)[0] if f.filename else f"Month {i + 1}"
                       for i, f in enumerate(files)]
        ws_values_list = [load_values_sheet(c) for c in contents]
        selections = company_breakdown_monthly(ws, ws_values_list, col_list, header_row)
    except (ValueError, TypeError, json.JSONDecodeError) as e:
        raise HTTPException(status_code=400, detail=str(e))

    rows, ad_rows, ad_names, ad_seen = [], [], [], set()
    pnl_keys = [
        ("revenue", "total_revenue"), ("cogs", "total_cogs"), ("gross_profit", "gross_profit"),
        ("operating_expense", "total_operating_expense"),
        ("net_operational_income", "net_operational_income"),
        ("other_income", "total_other_income"), ("other_expense", "total_other_expense"),
        ("net_other_income", "net_other_income"), ("net_income", "net_income"),
    ]

    for si, selection in enumerate(selections):
        company = selection_names[si]
        for mi, month in enumerate(selection["months"]):
            for metric_name, key in pnl_keys:
                pkr = float(month.get(key, 0) or 0)
                rows.append({
                    "company": company, "month": month_names[mi], "metric": metric_name,
                    "value": pkr / rate_list[mi], "pkr": pkr,
                    "label": f"{company} — {month_names[mi]} — {metric_name}",
                })
            if source == "cogs":
                # Use the structural COGS rows rather than cogs_items, which
                # intentionally omits zero/blank-value expenses. This keeps
                # the dropdown identical across companies with the same sheet.
                source_items = [
                    item for item in month.get("cogs_rows", [])
                    if item.get("kind") == "item"
                ]
            elif source == "opex":
                # OPEX = every Operating Expense item + every Other Expense item.
                opex_items = list(month.get("operating_expense_items", []))
                other_items = list(month.get("other_expense_items", []))
                opex_labels = {str(i.get("label", "")) for i in opex_items}
                source_items = opex_items + [
                    # same name in both groups -> keep them apart in the graph
                    {**i, "label": f"{i.get('label', '')} (Other Expense)"}
                    if str(i.get("label", "")) in opex_labels else i
                    for i in other_items
                ]
            else:
                source_items = month.get("revenue_items", [])
            default_name = {"cogs": "COGS item", "opex": "OPEX item"}.get(source, "Revenue source")
            for item in source_items:
                pkr = float(item.get("rs", 0) or 0)
                ad_name = str(item.get("label", default_name))
                if ad_name not in ad_seen:
                    ad_seen.add(ad_name)
                    ad_names.append(ad_name)
                ad_rows.append({
                    "company": company, "month": month_names[mi], "ad": ad_name,
                    "value": pkr / rate_list[mi], "pkr": pkr,
                    "label": f"{ad_name} — {company} — {month_names[mi]}",
                })

    try:
        requested_ads = json.loads(requested_ad_names) if requested_ad_names else []
    except json.JSONDecodeError:
        requested_ads = []
    if isinstance(requested_ads, list) and requested_ads:
        ad_rows = [r for r in ad_rows if r["ad"] in requested_ads]

    return {
        "rows": rows,
        "ad_rows": ad_rows,
        "ad_names": ad_names,
        "metrics": [x[0] for x in pnl_keys],
        "description": (
            "P&L values follow Revenue → COGS → Gross Profit → Operating "
            "Expense → Net Operational Income → Other Income/Expense → "
            "Net Profit/Loss. Each sheet is converted with its own PKR per USD rate."
        ),
    }


@router.post("/compare")
async def compare_route(files: List[UploadFile] = File(...), cols: str = Form(...)):
    """One comparison entry per selected entity (first file's layout)."""
    contents = [await f.read() for f in files]
    try:
        ws, total_col, header_row, last_col = build_reference(contents[:1])
        ws_values = load_values_sheet(contents[0])
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    try:
        col_list = [int(c) for c in cols.split(",") if c.strip() != ""]
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid column selection.")
    if not col_list or any(c < 1 or c > last_col for c in col_list):
        raise HTTPException(status_code=400, detail="Invalid column selection.")
    out = []
    for c in dict.fromkeys(col_list):
        r = company_breakdown(ws, ws_values, [c], header_row)
        out.append({"col": c, **r})
    return {"companies": out}