"""File download routes: Excel exports and PDFs."""
import io
import json
from typing import List, Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from .excel_export import export_selection
from .pdf_report import build_full_pdf, build_pdf_report, build_selection_pdf
from .rates import parse_rates
from .sheet_reader import find_groups_and_standalone, load_sheet
from .usd_column import add_usd_column

router = APIRouter()
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _month_list(months: str):
    return [m.strip() for m in months.split(",") if m.strip()]


@router.post("/export")
async def export_route(
    files: List[UploadFile] = File(...),
    rates: str = Form(""),
    rate: Optional[float] = Form(None),
    compare: int = Form(0),
    mode: str = Form(...),       # "group" | "single" | "grand" | "multi"
    key: str = Form(""),
    name: str = Form("Export"),
    months: str = Form(""),
):
    contents = [await f.read() for f in files]
    try:
        ws, _, total_col, header_row = load_sheet(contents[0])
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    groups, standalone = find_groups_and_standalone(ws, total_col)

    if mode == "grand":
        cols = [total_col]
    elif mode == "group":
        g = groups.get(key)
        if not g:
            raise HTTPException(status_code=400, detail=f'Group "{key}" not found.')
        cols = [m["col"] for m in g["members"]]
        if g["total_col"]:
            cols.append(g["total_col"])
    elif mode == "single":
        try:
            col = int(key)
        except ValueError:
            raise HTTPException(status_code=400, detail="Invalid column key.")
        if col < 1 or col > total_col:
            raise HTTPException(status_code=400, detail="Invalid company column.")
        cols = [col]
    elif mode == "multi":
        try:
            cols = [int(c) for c in key.split(",") if c.strip() != ""]
        except ValueError:
            raise HTTPException(status_code=400, detail="Invalid column selection.")
        if not cols:
            raise HTTPException(status_code=400, detail="No columns selected.")
        for c in cols:
            if c < 1 or c > total_col:
                raise HTTPException(status_code=400, detail="Invalid company column.")
        cols = sorted(set(cols))
    else:
        raise HTTPException(status_code=400, detail="mode must be group, single, grand, or multi.")

    try:
        rate_list = parse_rates(rates, len(contents), rate)
        if compare and len(contents) < 2:
            raise ValueError("Comparison needs at least 2 months.")
        out_bytes = export_selection(contents, rate_list, cols, name,
                                     month_names=_month_list(months), compare=bool(compare))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    safe_name = "".join(ch if ch.isalnum() or ch in " _-." else "_" for ch in name).strip() or "Export"
    filename = f"{safe_name}_Comparison_USD.xlsx" if compare else f"{safe_name}_USD.xlsx"
    return StreamingResponse(io.BytesIO(out_bytes), media_type=XLSX,
                             headers={"Content-Disposition": f'attachment; filename="{filename}"'})


@router.post("/export_pdf")
async def export_pdf_route(
    files: List[UploadFile] = File(...),
    rates: str = Form(""),
    rate: Optional[float] = Form(None),
    key: str = Form(...),
    names: str = Form(""),
    months: str = Form(""),
):
    contents = [await f.read() for f in files]
    try:
        ws, _, total_col, _ = load_sheet(contents[0])
        cols = sorted({int(c) for c in key.split(",") if c.strip()})
    except (ValueError, TypeError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    if not cols or any(c < 1 or c > total_col for c in cols):
        raise HTTPException(status_code=400, detail="Invalid column selection.")
    try:
        name_list = json.loads(names) if names else []
    except json.JSONDecodeError:
        name_list = []
    # names arrive in tick order; cols are sorted, so re-pair them by column.
    try:
        orig_cols = [int(c) for c in key.split(",") if c.strip()]
        pair = dict(zip(orig_cols, name_list)) if len(name_list) == len(orig_cols) else {}
    except ValueError:
        pair = {}
    name_list = [pair.get(c, f"Column {c}") for c in cols]
    try:
        pdf_bytes = build_selection_pdf(contents, parse_rates(rates, len(contents), rate),
                                        cols, name_list, _month_list(months))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    return StreamingResponse(io.BytesIO(pdf_bytes), media_type="application/pdf",
                             headers={"Content-Disposition": 'attachment; filename="PnL_Statement.pdf"'})


@router.post("/export_full_pdf")
async def export_full_pdf_route(
    files: List[UploadFile] = File(...),
    rates: str = Form(""),
    rate: Optional[float] = Form(None),
    key: str = Form(...),
    names: str = Form(""),
    months: str = Form(""),
):
    """ONE PDF = P&L statement tables + graphs/revenue pages."""
    contents = [await f.read() for f in files]
    try:
        ws, _, total_col, _ = load_sheet(contents[0])
        orig_cols = [int(c) for c in key.split(",") if c.strip()]
    except (ValueError, TypeError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    if not orig_cols or any(c < 1 or c > total_col for c in orig_cols):
        raise HTTPException(status_code=400, detail="Invalid column selection.")
    try:
        rate_list = parse_rates(rates, len(contents), rate)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    try:
        name_list = json.loads(names) if names else []
    except json.JSONDecodeError:
        name_list = []
    if not isinstance(name_list, list) or len(name_list) != len(orig_cols):
        name_list = [f"Column {c}" for c in orig_cols]
    try:
        pdf_bytes = build_full_pdf(contents, rate_list, orig_cols, name_list, _month_list(months))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    return StreamingResponse(io.BytesIO(pdf_bytes), media_type="application/pdf",
                             headers={"Content-Disposition": 'attachment; filename="PnL_Full_Report.pdf"'})


@router.post("/report_pdf")
async def report_pdf_route(
    files: List[UploadFile] = File(...),
    rates: str = Form(""),
    rate: Optional[float] = Form(None),
    cols: str = Form(...),
    names: str = Form(""),
):
    contents = [await f.read() for f in files]
    if not contents:
        raise HTTPException(status_code=400, detail="No files selected.")
    try:
        ws, _, total_col, _ = load_sheet(contents[0])
        col_list = [int(c) for c in cols.split(",") if c.strip()]
    except (ValueError, TypeError):
        raise HTTPException(status_code=400, detail="Invalid column selection.")
    if not col_list or any(c < 1 or c > total_col for c in col_list):
        raise HTTPException(status_code=400, detail="Invalid column selection.")
    try:
        selection_names = json.loads(names) if names else []
    except json.JSONDecodeError:
        selection_names = []
    if not isinstance(selection_names, list) or len(selection_names) != len(col_list):
        selection_names = [f"Column {c}" for c in col_list]
    try:
        month_names = [f.filename.rsplit(".", 1)[0] if f.filename else f"Month {i + 1}"
                       for i, f in enumerate(files)]
        pdf_bytes = build_pdf_report(contents, parse_rates(rates, len(contents), rate),
                                     col_list, selection_names, month_names=month_names)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    return StreamingResponse(io.BytesIO(pdf_bytes), media_type="application/pdf",
                             headers={"Content-Disposition": 'attachment; filename="Revenue_Report.pdf"'})


@router.post("/process")
async def process(file: UploadFile = File(...), rate: float = Form(...)):
    content = await file.read()
    try:
        out_bytes, _ = add_usd_column(content, rate)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    filename = file.filename.rsplit(".", 1)[0] + "_with_USD.xlsx"
    return StreamingResponse(io.BytesIO(out_bytes), media_type=XLSX,
                             headers={"Content-Disposition": f'attachment; filename="{filename}"'})
