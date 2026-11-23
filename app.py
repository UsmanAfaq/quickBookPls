"""
Rs -> USD Ledger Tool — FastAPI entry point.

All logic lives in the `ledger/` package (one small file per job):
    config.py          settings, folders, secret
    storage.py         sheets.json read/write
    auth.py            login, session cookie, access guard
    routes_sheets.py   /api/sheets*, dashboard page
    routes_data.py     JSON routes (structure, breakdown, graph data, comparison)
    routes_downloads.py Excel / PDF download routes
    sheet_reader.py    detect TOTAL column, companies, groups, labels
    rates.py           exchange-rate helpers, month pairs
    pnl.py             P&L hierarchy + USD conversion
    usd_column.py      single-file "add USD column"
    excel_export.py    filtered / comparison Excel workbook
    pdf_report.py      charts + PDF reports
    templates/         login.html, dashboard.html (front end)

Run:  uvicorn app:app --reload   ->  http://127.0.0.1:8000
"""
from fastapi import FastAPI

from ledger import auth, routes_data, routes_downloads, routes_sheets, quickbooks_desktop

app = FastAPI(title="Rs to USD Ledger Tool")
app.middleware("http")(auth.auth_guard)

app.include_router(auth.router)
app.include_router(routes_sheets.router)
app.include_router(routes_data.router)
app.include_router(routes_downloads.router)
app.include_router(quickbooks_desktop.router)
