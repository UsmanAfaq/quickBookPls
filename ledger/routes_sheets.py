# """Server-side sheet storage routes (/api/sheets...) and the dashboard page."""
# import time
# import uuid
# from pathlib import Path
# from typing import List, Optional
#
# from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
# from fastapi.responses import HTMLResponse, Response
#
# from .config import META_LOCK, SHEETS_DIR
# from .sheet_reader import ensure_total_column, load_sheet
# from .storage import load_companies, load_meta, new_company, save_companies, save_meta
#
# router = APIRouter()
# DASHBOARD_FILE = Path(__file__).parent / "templates" / "dashboard.html"
#
#
# @router.get("/", response_class=HTMLResponse)
# def home(request: Request):
#     # read on every request, so an updated dashboard.html shows up without restarting the server
#     page = DASHBOARD_FILE.read_text(encoding="utf-8").replace("__ROLE__", request.state.role)
#     return HTMLResponse(page, headers={"Cache-Control": "no-store"})
#
#
# def _company_or_404(companies, company_id, role):
#     """The company, or 404. Guests may only see ENABLED companies."""
#     c = next((x for x in companies if x["id"] == company_id), None)
#     if c is None or (role != "admin" and not c.get("enabled", True)):
#         raise HTTPException(status_code=404, detail="Company not found.")
#     return c
#
#
# # ---- companies (workspaces) -------------------------------------------------
# @router.get("/api/companies")
# def api_list_companies(request: Request):
#     """Admin sees all (with enabled flag); guests see only enabled companies."""
#     with META_LOCK:
#         companies = load_companies()
#         counts = {}
#         for m in load_meta():
#             counts[m.get("company_id")] = counts.get(m.get("company_id"), 0) + 1
#     if request.state.role != "admin":
#         companies = [c for c in companies if c.get("enabled", True)]
#     return {"companies": [{**c, "sheet_count": counts.get(c["id"], 0)} for c in companies]}
#
#
# @router.post("/api/companies")
# def api_create_company(name: str = Form(...)):
#     name = name.strip()
#     if not name or len(name) > 80:
#         raise HTTPException(status_code=400, detail="Enter a company name (max 80 characters).")
#     with META_LOCK:
#         companies = load_companies()
#         if any(c["name"].lower() == name.lower() for c in companies):
#             raise HTTPException(status_code=400, detail=f'A company named "{name}" already exists.')
#         c = new_company(name)
#         save_companies(companies + [c])
#     return {"company": {**c, "sheet_count": 0}}
#
#
# @router.put("/api/companies/{company_id}")
# def api_update_company(company_id: str, name: Optional[str] = Form(None), enabled: Optional[bool] = Form(None)):
#     """Administrator: rename and/or enable / disable a company."""
#     with META_LOCK:
#         companies = load_companies()
#         c = next((x for x in companies if x["id"] == company_id), None)
#         if c is None:
#             raise HTTPException(status_code=404, detail="Company not found.")
#         if name is not None:
#             name = name.strip()
#             if not name or len(name) > 80:
#                 raise HTTPException(status_code=400, detail="Enter a company name (max 80 characters).")
#             if any(x["id"] != company_id and x["name"].lower() == name.lower() for x in companies):
#                 raise HTTPException(status_code=400, detail=f'A company named "{name}" already exists.')
#             c["name"] = name
#         if enabled is not None:
#             c["enabled"] = enabled
#         save_companies(companies)
#     return {"company": c}
#
#
# @router.delete("/api/companies/{company_id}")
# def api_delete_company(company_id: str):
#     """Administrator: deletes the company AND all of its sheets."""
#     with META_LOCK:
#         companies = load_companies()
#         if not any(x["id"] == company_id for x in companies):
#             raise HTTPException(status_code=404, detail="Company not found.")
#         meta = load_meta()
#         gone = [m for m in meta if m.get("company_id") == company_id]
#         save_meta([m for m in meta if m.get("company_id") != company_id])
#         save_companies([x for x in companies if x["id"] != company_id])
#     for m in gone:
#         try:
#             (SHEETS_DIR / f"{m['id']}.xlsx").unlink()
#         except OSError:
#             pass
#     return {"ok": True, "removed_sheets": len(gone)}
#
#
# # ---- sheets (always inside one company) -------------------------------------
# @router.get("/api/sheets")
# def api_list_sheets(request: Request, company_id: Optional[str] = None):
#     with META_LOCK:
#         companies = load_companies()
#         if company_id is None and companies:          # old clients: first enabled company
#             first = next((c for c in companies if c.get("enabled", True)), companies[0])
#             company_id = first["id"]
#         _company_or_404(companies, company_id, request.state.role)
#         return {"sheets": [m for m in load_meta() if m.get("company_id") == company_id]}
#
#
# @router.post("/api/sheets")
# async def api_upload_sheets(files: List[UploadFile] = File(...), rate: Optional[float] = Form(None),
#                             company_id: str = Form(...)):
#     """Administrator only. Saves each .xlsx on the server (so guests can view it)."""
#     with META_LOCK:
#         _company_or_404(load_companies(), company_id, "admin")
#     if rate is not None and rate <= 0:
#         rate = None
#     added = []
#     for f in files:
#         name = f.filename or "sheet.xlsx"
#         if not name.lower().endswith(".xlsx"):
#             raise HTTPException(status_code=400, detail=f"{name}: only .xlsx files are allowed.")
#         content = await f.read()
#         try:
#             content = ensure_total_column(content)  # no TOTAL column -> calculate one
#             load_sheet(content)
#         except Exception:
#             raise HTTPException(status_code=400, detail=f"{name}: this is not a valid P&L export (no TOTAL header found).")
#         sid = uuid.uuid4().hex
#         (SHEETS_DIR / f"{sid}.xlsx").write_bytes(content)
#         added.append({"id": sid, "name": name[:-5], "filename": name, "rate": rate,
#                       "company_id": company_id, "uploaded_at": int(time.time())})
#     with META_LOCK:
#         save_meta(load_meta() + added)
#     return {"sheets": added}
#
#
# @router.delete("/api/sheets/{sheet_id}")
# def api_delete_sheet(sheet_id: str):
#     with META_LOCK:
#         meta = load_meta()
#         if not any(m["id"] == sheet_id for m in meta):
#             raise HTTPException(status_code=404, detail="Sheet not found.")
#         save_meta([m for m in meta if m["id"] != sheet_id])
#     try:
#         (SHEETS_DIR / f"{sheet_id}.xlsx").unlink()
#     except OSError:
#         pass
#     return {"ok": True}
#
#
# @router.put("/api/sheets/{sheet_id}/rate")
# def api_set_rate(sheet_id: str, rate: float = Form(...)):
#     """Administrator only: this sheet's own PKR-per-USD rate."""
#     if rate <= 0:
#         raise HTTPException(status_code=400, detail="Rate must be greater than 0.")
#     with META_LOCK:
#         meta = load_meta()
#         for m in meta:
#             if m["id"] == sheet_id:
#                 m["rate"] = rate
#                 save_meta(meta)
#                 return {"id": sheet_id, "rate": rate}
#     raise HTTPException(status_code=404, detail="Sheet not found.")
#
#
# @router.get("/api/sheets/{sheet_id}/file")
# def api_sheet_file(sheet_id: str, request: Request):
#     if not sheet_id.isalnum():
#         raise HTTPException(status_code=404, detail="Sheet not found.")
#     with META_LOCK:
#         companies = load_companies()
#         m = next((x for x in load_meta() if x["id"] == sheet_id), None)
#     if m is None:
#         raise HTTPException(status_code=404, detail="Sheet not found.")
#     _company_or_404(companies, m.get("company_id"), request.state.role)
#     f = SHEETS_DIR / f"{sheet_id}.xlsx"
#     if not f.exists():
#         raise HTTPException(status_code=404, detail="Sheet not found.")
#     return Response(f.read_bytes(),
#                     media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
"""Server-side sheet storage routes (/api/sheets...) and the dashboard page."""
import json
import time
import uuid
from pathlib import Path
from typing import List, Optional

from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response

from .config import META_LOCK, SHEETS_DIR
from .sheet_reader import ensure_total_column, load_sheet
from .storage import load_companies, load_meta, new_company, save_companies, save_meta

router = APIRouter()
DASHBOARD_FILE = Path(__file__).parent / "templates" / "dashboard.html"


@router.get("/", response_class=HTMLResponse)
def home(request: Request):
    # read on every request, so an updated dashboard.html shows up without restarting the server
    page = (DASHBOARD_FILE.read_text(encoding="utf-8")
            .replace("__ROLE__", request.state.role)
            .replace("__PERMS__", json.dumps(request.state.perms)))
    return HTMLResponse(page, headers={"Cache-Control": "no-store"})


def _same_name(a, b):
    return " ".join(str(a).lower().split()) == " ".join(str(b).lower().split())


def _can_see(company, role, scope=None):
    """Admin: everything. Guest: every ENABLED company. Accountant: only their own company."""
    if role == "admin":
        return True
    if not company.get("enabled", True):
        return False
    return role != "accountant" or (scope is not None and _same_name(company["name"], scope))


def _company_or_404(companies, company_id, role, scope=None):
    """The company, or 404 when it does not exist or this login may not see it."""
    c = next((x for x in companies if x["id"] == company_id), None)
    if c is None or not _can_see(c, role, scope):
        raise HTTPException(status_code=404, detail="Company not found.")
    return c


# ---- companies (workspaces) -------------------------------------------------
@router.get("/api/companies")
def api_list_companies(request: Request):
    """Admin sees all (with enabled flag); guests see only enabled companies."""
    with META_LOCK:
        companies = load_companies()
        counts = {}
        for m in load_meta():
            counts[m.get("company_id")] = counts.get(m.get("company_id"), 0) + 1
    companies = [c for c in companies if _can_see(c, request.state.role, request.state.scope)]
    return {"companies": [{**c, "sheet_count": counts.get(c["id"], 0)} for c in companies]}


@router.post("/api/companies")
def api_create_company(name: str = Form(...)):
    name = name.strip()
    if not name or len(name) > 80:
        raise HTTPException(status_code=400, detail="Enter a company name (max 80 characters).")
    with META_LOCK:
        companies = load_companies()
        if any(c["name"].lower() == name.lower() for c in companies):
            raise HTTPException(status_code=400, detail=f'A company named "{name}" already exists.')
        c = new_company(name)
        save_companies(companies + [c])
    return {"company": {**c, "sheet_count": 0}}


@router.put("/api/companies/{company_id}")
def api_update_company(company_id: str, name: Optional[str] = Form(None), enabled: Optional[bool] = Form(None)):
    """Administrator: rename and/or enable / disable a company."""
    with META_LOCK:
        companies = load_companies()
        c = next((x for x in companies if x["id"] == company_id), None)
        if c is None:
            raise HTTPException(status_code=404, detail="Company not found.")
        if name is not None:
            name = name.strip()
            if not name or len(name) > 80:
                raise HTTPException(status_code=400, detail="Enter a company name (max 80 characters).")
            if any(x["id"] != company_id and x["name"].lower() == name.lower() for x in companies):
                raise HTTPException(status_code=400, detail=f'A company named "{name}" already exists.')
            c["name"] = name
        if enabled is not None:
            c["enabled"] = enabled
        save_companies(companies)
    return {"company": c}


@router.delete("/api/companies/{company_id}")
def api_delete_company(company_id: str):
    """Administrator: deletes the company AND all of its sheets."""
    with META_LOCK:
        companies = load_companies()
        if not any(x["id"] == company_id for x in companies):
            raise HTTPException(status_code=404, detail="Company not found.")
        meta = load_meta()
        gone = [m for m in meta if m.get("company_id") == company_id]
        save_meta([m for m in meta if m.get("company_id") != company_id])
        save_companies([x for x in companies if x["id"] != company_id])
    for m in gone:
        try:
            (SHEETS_DIR / f"{m['id']}.xlsx").unlink()
        except OSError:
            pass
    return {"ok": True, "removed_sheets": len(gone)}


# ---- sheets (always inside one company) -------------------------------------
@router.get("/api/sheets")
def api_list_sheets(request: Request, company_id: Optional[str] = None):
    with META_LOCK:
        companies = load_companies()
        if company_id is None and companies:          # old clients: first company this login may see
            allowed = [c for c in companies if _can_see(c, request.state.role, request.state.scope)]
            first = next((c for c in allowed if c.get("enabled", True)), allowed[0] if allowed else None)
            company_id = first["id"] if first else None
        _company_or_404(companies, company_id, request.state.role, request.state.scope)
        return {"sheets": [m for m in load_meta() if m.get("company_id") == company_id]}


def _require_uploader(request: Request, company: dict) -> None:
    """Administrator: any company. Accountant: ONLY their own company. Guest: never."""
    role = request.state.role
    if role == "admin":
        return
    if role == "accountant" and _same_name(company["name"], request.state.scope or ""):
        return
    raise HTTPException(status_code=403, detail="You can only upload sheets to your own company.")


@router.post("/api/sheets")
async def api_upload_sheets(request: Request, files: List[UploadFile] = File(...),
                            rate: Optional[float] = Form(None), company_id: str = Form(...)):
    """Administrator (any company) or accountant (own company only). Saves each .xlsx on the server."""
    with META_LOCK:
        company = _company_or_404(load_companies(), company_id, "admin")
    _require_uploader(request, company)
    if rate is not None and rate <= 0:
        rate = None
    added = []
    for f in files:
        name = f.filename or "sheet.xlsx"
        if not name.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"{name}: only .xlsx files are allowed.")
        content = await f.read()
        try:
            content = ensure_total_column(content)  # no TOTAL column -> calculate one
            load_sheet(content)
        except Exception:
            raise HTTPException(status_code=400, detail=f"{name}: this is not a valid P&L export (no TOTAL header found).")
        sid = uuid.uuid4().hex
        (SHEETS_DIR / f"{sid}.xlsx").write_bytes(content)
        added.append({"id": sid, "name": name[:-5], "filename": name, "rate": rate,
                      "company_id": company_id, "uploaded_at": int(time.time())})
    with META_LOCK:
        save_meta(load_meta() + added)
    return {"sheets": added}


@router.delete("/api/sheets/{sheet_id}")
def api_delete_sheet(request: Request, sheet_id: str):
    """Administrator (any sheet) or an accountant with the Edit/Delete permission (own company only)."""
    with META_LOCK:
        companies = load_companies()
        meta = load_meta()
        m = next((x for x in meta if x["id"] == sheet_id), None)
        if m is None:
            raise HTTPException(status_code=404, detail="Sheet not found.")
        company = next((c for c in companies if c["id"] == m.get("company_id")), None)
        if company is None:
            raise HTTPException(status_code=404, detail="Sheet not found.")
        _require_uploader(request, company)
        save_meta([m for m in meta if m["id"] != sheet_id])
    try:
        (SHEETS_DIR / f"{sheet_id}.xlsx").unlink()
    except OSError:
        pass
    return {"ok": True}


@router.put("/api/sheets/{sheet_id}/rate")
def api_set_rate(request: Request, sheet_id: str, rate: float = Form(...)):
    """Administrator (any sheet) or accountant (sheets of their own company): this sheet's PKR-per-USD rate."""
    if rate <= 0:
        raise HTTPException(status_code=400, detail="Rate must be greater than 0.")
    with META_LOCK:
        companies = load_companies()
        meta = load_meta()
        for m in meta:
            if m["id"] == sheet_id:
                company = next((c for c in companies if c["id"] == m.get("company_id")), None)
                if company is None:
                    raise HTTPException(status_code=404, detail="Sheet not found.")
                _require_uploader(request, company)
                m["rate"] = rate
                save_meta(meta)
                return {"id": sheet_id, "rate": rate}
    raise HTTPException(status_code=404, detail="Sheet not found.")


@router.get("/api/sheets/{sheet_id}/file")
def api_sheet_file(sheet_id: str, request: Request):
    if not sheet_id.isalnum():
        raise HTTPException(status_code=404, detail="Sheet not found.")
    with META_LOCK:
        companies = load_companies()
        m = next((x for x in load_meta() if x["id"] == sheet_id), None)
    if m is None:
        raise HTTPException(status_code=404, detail="Sheet not found.")
    _company_or_404(companies, m.get("company_id"), request.state.role, request.state.scope)
    f = SHEETS_DIR / f"{sheet_id}.xlsx"
    if not f.exists():
        raise HTTPException(status_code=404, detail="Sheet not found.")
    return Response(f.read_bytes(),
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")