"""QuickBooks Desktop Web Connector bridge for PLS Accounts.

QBWC polls this HTTPS SOAP endpoint. Credentials are supplied via environment:
QBWC_USERNAME, QBWC_PASSWORD, PLS_PUBLIC_URL. A dashboard admin queues a P&L
report using POST /api/quickbooks/sync with company_id, start_date, end_date.
"""
import html
import io
import json
import os
import secrets
import time
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import Response
from openpyxl import Workbook

from .config import DATA_DIR, META_LOCK, SHEETS_DIR
from .sheet_reader import ensure_total_column, load_sheet
from .storage import load_companies, load_meta, save_meta

router = APIRouter()
QBWC_STATE = DATA_DIR / "quickbooks_desktop.json"
QBWC_USERNAME = os.getenv("QBWC_USERNAME", "pls")
QBWC_PASSWORD = os.getenv("QBWC_PASSWORD", "")
PUBLIC_URL = os.getenv("PLS_PUBLIC_URL", "").rstrip("/")
OWNER_ID = "{" + str(uuid.uuid5(uuid.NAMESPACE_DNS, "pls-accounts.quickbooks.owner")) + "}"
FILE_ID = "{" + str(uuid.uuid5(uuid.NAMESPACE_DNS, "pls-accounts.quickbooks.file")) + "}"
SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
QBWC_NS = "http://developer.intuit.com/"


def _state():
    try:
        return json.loads(QBWC_STATE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"pending": None, "active": None, "last_sync": None, "last_error": None}


def _save_state(s):
    tmp = QBWC_STATE.with_suffix(".tmp")
    tmp.write_text(json.dumps(s, indent=2), encoding="utf-8")
    tmp.replace(QBWC_STATE)


def _soap_response(method: str, values):
    # QBWC expects a SOAP 1.1 response whose result is an array of strings or a scalar.
    if isinstance(values, (list, tuple)):
        result = "".join(f"<string>{html.escape(str(v))}</string>" for v in values)
        result = f"<{method}Result>{result}</{method}Result>"
    else:
        result = f"<{method}Result>{html.escape(str(values))}</{method}Result>"
    body = (f'<?xml version="1.0" encoding="utf-8"?>'
            f'<soap:Envelope xmlns:soap="{SOAP_NS}" xmlns="{QBWC_NS}">'
            f'<soap:Body><{method}Response>{result}</{method}Response></soap:Body></soap:Envelope>')
    return Response(body, media_type="text/xml; charset=utf-8")


def _text(root, name, default=""):
    for el in root.iter():
        if el.tag.split("}")[-1] == name:
            return el.text or default
    return default


def _qbxml_request(job):
    return ("<?xml version=\"1.0\"?><?qbxml version=\"13.0\"?>"
            "<QBXML><QBXMLMsgsRq onError=\"stopOnError\">"
            "<GeneralSummaryReportQueryRq requestID=\"1\">"
            "<GeneralSummaryReportType>ProfitAndLossStandard</GeneralSummaryReportType>"
            f"<ReportPeriod><FromReportDate>{job['start_date']}</FromReportDate>"
            f"<ToReportDate>{job['end_date']}</ToReportDate></ReportPeriod>"
            "<SummarizeColumnsBy>TotalOnly</SummarizeColumnsBy><ReturnRows>All</ReturnRows>"
            "<ReportBasis>Accrual</ReportBasis>"
            "</GeneralSummaryReportQueryRq></QBXMLMsgsRq></QBXML>")


def _report_to_xlsx(response_xml: str, company_name: str, period_name: str) -> bytes:
    root = ET.fromstring(response_xml)
    report = next((e for e in root.iter() if e.tag.split("}")[-1] == "ReportRet"), None)
    if report is None:
        # QBXML error responses must not be silently imported as empty reports.
        err = next((e for e in root.iter() if e.tag.split("}")[-1] == "StatusMessage"), None)
        raise ValueError(err.text if err is not None and err.text else "QuickBooks returned no Profit & Loss report.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Profit and Loss"
    ws.cell(1, 1, None)
    ws.cell(1, 2, company_name)
    ws.cell(2, 3, "TOTAL")
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20

    row_num = 3
    # Report rows are hierarchical. Preserve the displayed account/section labels and
    # the numeric value in the first report column; this sync is intentionally one month.
    def walk(parent):
        nonlocal row_num
        for child in list(parent):
            tag = child.tag.split("}")[-1]
            if tag in ("DataRow", "TextRow", "SubtotalRow", "TotalRow"):
                cols = [x for x in child if x.tag.split("}")[-1] == "ColData"]
                row_data = next(
                    (x for x in child if x.tag.split("}")[-1] == "RowData"),
                    None
                )

                label = ""
                if row_data is not None:
                    label = row_data.attrib.get("value", "").strip()
                if not label and cols:
                    label = cols[0].attrib.get("value", "").strip()

                if label:
                    ws.cell(row_num, 1, label)

                    amount = None
                    for col in cols[1:]:
                        raw = col.attrib.get("value", "").replace(",", "").replace("$", "").strip()
                        try:
                            amount = float(raw) if raw else amount
                        except ValueError:
                            pass

                    if amount is not None:
                        ws.cell(row_num, 2, amount)
                        ws.cell(row_num, 3, amount)

                    row_num += 1
            walk(child)
    walk(report)
    if row_num == 3:
        raise ValueError("QuickBooks P&L report contained no account rows to import.")
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@router.post("/api/quickbooks/sync")
async def queue_quickbooks_sync(request: Request):
    """Queue a single-period P&L import. Dashboard session/admin required."""
    if request.state.role != "admin":
        raise HTTPException(status_code=403, detail="Only the dashboard administrator can start QuickBooks sync.")
    if not QBWC_PASSWORD:
        raise HTTPException(status_code=503, detail="Set QBWC_PASSWORD on the VPS before enabling QuickBooks sync.")
    payload = await request.json()
    company_id = str(payload.get("company_id", ""))
    start_date = str(payload.get("start_date", ""))
    end_date = str(payload.get("end_date", ""))
    try:
        time.strptime(start_date, "%Y-%m-%d")
        time.strptime(end_date, "%Y-%m-%d")
    except ValueError:
        raise HTTPException(status_code=400, detail="Dates must use YYYY-MM-DD format.")
    if end_date < start_date:
        raise HTTPException(status_code=400, detail="End date must not be before start date.")
    if start_date[:7] != end_date[:7]:
        raise HTTPException(status_code=400, detail="Sync one calendar month at a time so PLS can store each month separately.")
    with META_LOCK:
        companies = load_companies()
        company = next((c for c in companies if c["id"] == company_id), None)
    if not company:
        raise HTTPException(status_code=404, detail="PLS company not found.")
    state = _state()
    if state.get("pending") or state.get("active"):
        raise HTTPException(status_code=409, detail="A QuickBooks sync is already queued or running.")
    state["pending"] = {"company_id": company_id, "company_name": company["name"],
                         "start_date": start_date, "end_date": end_date,
                         "period_name": start_date[:7], "queued_at": int(time.time())}
    state["last_error"] = None
    _save_state(state)
    return {"ok": True, "message": "P&L sync queued. Select PLS Accounts in QuickBooks Web Connector and click Update Selected."}


@router.get("/api/quickbooks/status")
def quickbooks_status(request: Request):
    if request.state.role != "admin":
        raise HTTPException(status_code=403, detail="Administrator access required.")
    s = _state()
    return {"configured": bool(QBWC_PASSWORD and PUBLIC_URL), "queued": bool(s.get("pending")),
            "running": bool(s.get("active")), "last_sync": s.get("last_sync"),
            "last_error": s.get("last_error"), "public_url_configured": bool(PUBLIC_URL)}


@router.get("/quickbooks/pls.qwc")
def download_qwc(request: Request):
    # This endpoint is behind the normal dashboard login; it must not be public.
    if not PUBLIC_URL or not QBWC_PASSWORD:
        raise HTTPException(status_code=503, detail="Configure PLS_PUBLIC_URL and QBWC_PASSWORD on the VPS first.")
    # QWC download is performed from a logged-in browser. auth_guard handles the login.
    xml = f'''<?xml version="1.0"?>
<QBWCXML>
  <AppName>PLS Accounts</AppName><AppID></AppID><AppURL>{html.escape(PUBLIC_URL + '/qbwc')}</AppURL>
  <AppDescription>Import QuickBooks Desktop Profit and Loss reports into PLS Accounts.</AppDescription>
  <AppSupport>https://developer.intuit.com/</AppSupport><UserName>{html.escape(QBWC_USERNAME)}</UserName>
  <OwnerID>{OWNER_ID}</OwnerID><FileID>{FILE_ID}</FileID><QBType>QBFS</QBType>
  <Scheduler><RunEveryNMinutes>60</RunEveryNMinutes></Scheduler>
  <IsReadOnly>true</IsReadOnly><PersonalDataPref>pdpNotNeeded</PersonalDataPref>
</QBWCXML>'''
    return Response(xml, media_type="application/xml", headers={"Content-Disposition": 'attachment; filename="PLS_Accounts.qwc"'})


@router.post("/qbwc")
async def qbwc_soap(request: Request):
    raw = await request.body()
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        raise HTTPException(status_code=400, detail="Invalid SOAP XML.")
    body = next((e for e in root.iter() if e.tag.split("}")[-1] == "Body"), None)
    method_el = list(body)[0] if body is not None and list(body) else None
    if method_el is None:
        raise HTTPException(status_code=400, detail="Missing SOAP method.")
    method = method_el.tag.split("}")[-1]
    args = {e.tag.split("}")[-1]: (e.text or "") for e in method_el.iter() if e is not method_el}
    state = _state()

    if method == "authenticate":
        if args.get("strUserName") != QBWC_USERNAME or not secrets.compare_digest(args.get("strPassword", ""), QBWC_PASSWORD):
            return _soap_response(method, ["", ""])
        if not state.get("pending"):
            return _soap_response(method, ["", ""])
        ticket = uuid.uuid4().hex
        state["active"] = {**state["pending"], "ticket": ticket}
        state["pending"] = None
        _save_state(state)
        return _soap_response(method, [ticket, ""])
    if method == "sendRequestXML":
        active = state.get("active")
        return _soap_response(method, _qbxml_request(active) if active else "")
    if method == "receiveResponseXML":
        active = state.get("active")
        response_xml = args.get("response", "")
        if response_xml:
            debug_path = Path(__file__).resolve().parent.parent / "data" / "qb_response_debug.xml"
            debug_path.write_text(response_xml, encoding="utf-8")
        if not active:
            return _soap_response(method, 100)
        try:
            content = _report_to_xlsx(response_xml, active["company_name"], active["period_name"])
            content = ensure_total_column(content)
            load_sheet(content)
            sid = uuid.uuid4().hex
            filename = f"QuickBooks P&L {active['period_name']}.xlsx"
            (SHEETS_DIR / f"{sid}.xlsx").write_bytes(content)
            with META_LOCK:
                meta = load_meta()
                # Replace prior QB-imported sheet for the same company and period; manual sheets remain untouched.
                old = [m for m in meta if m.get("source") == "quickbooks_desktop" and m.get("company_id") == active["company_id"] and m.get("period") == active["period_name"]]
                meta = [m for m in meta if m not in old]
                for m in old:
                    try: (SHEETS_DIR / f"{m['id']}.xlsx").unlink()
                    except OSError: pass
                meta.append({"id": sid, "name": f"QuickBooks P&L {active['period_name']}", "filename": filename,
                             "rate": None, "company_id": active["company_id"], "uploaded_at": int(time.time()),
                             "source": "quickbooks_desktop", "period": active["period_name"]})
                save_meta(meta)
            state["last_sync"] = {"at": int(time.time()), "company": active["company_name"], "period": active["period_name"]}
            state["last_error"] = None
        except Exception as exc:
            state["last_error"] = str(exc)[:500]
        state["active"] = None
        _save_state(state)
        return _soap_response(method, 100)
    if method == "getLastError":
        return _soap_response(method, state.get("last_error") or "No error")
    if method in ("clientVersion", "getClientVersion", "serverVersion", "getServerVersion"):
        return _soap_response(method, "")
    if method in ("closeConnection", "connectionError", "getInteractiveURL", "interactiveDone", "interactiveRejected"):
        if method == "closeConnection":
            state["active"] = None
            _save_state(state)
        return _soap_response(method, "")
    raise HTTPException(status_code=400, detail=f"Unsupported QBWC method: {method}")
