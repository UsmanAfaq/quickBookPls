
"""Login (Administrator / Accountant / Guest), signed session cookie, and the access-guard middleware."""
import base64
import hmac
import hashlib
import time
import json
import secrets

from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse

from .config import (ACCOUNTANTS, ADMIN_PASSWORD, ADMIN_USERNAME, GUEST_PASSWORD, SECRET,
                     SESSION_COOKIE, SESSION_SECONDS, DATA_DIR, META_LOCK)
from .storage import load_companies

router = APIRouter()
LOGIN_PAGE = (Path(__file__).parent / "templates" / "login.html").read_text(encoding="utf-8")

# Only the administrator may remove sheets or manage companies.
# (Uploading a sheet and setting its rate -- POST /api/sheets, PUT /api/sheets/{id}/rate -- are also
#  open to an accountant, but ONLY for their own company; routes_sheets.py enforces that.
#  Guests are blocked from them by the check in auth_guard below.)
ADMIN_ONLY = [("POST", "/api/companies"), ("PUT", "/api/companies/"), ("DELETE", "/api/companies/")]
# Permission-controlled routes (set per user in User Management):
#   view     -> can open sheets / P&L / graphs / comparisons
#   upload   -> can upload new Excel sheets (and set the USD rate on them)
#   download -> can download Excel / PDF reports
#   edit     -> can edit rates and delete sheets
PERMISSION_KEYS = ("view", "upload", "download", "edit")
DOWNLOAD_PATHS = ("/export", "/export_pdf", "/export_full_pdf", "/report_pdf", "/process")
ROLE_DEFAULT_PERMS = {
    "accountant": {"view": True, "upload": True, "download": True, "edit": False},
    "guest": {"view": True, "upload": False, "download": True, "edit": False},
}
PUBLIC_PATHS = ("/login", "/logout")
USERS_FILE = DATA_DIR / "users.json"


def _load_users():
    """Load managed non-admin accounts; seed legacy accountant logins once."""
    try:
        users = json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        users = {}
    changed = False
    for username, info in ACCOUNTANTS.items():
        if username not in users:
            salt = secrets.token_hex(16)
            digest = hashlib.pbkdf2_hmac("sha256", info["password"].encode(), salt.encode(), 180000).hex()
            users[username] = {"role": "accountant", "company": info["company"], "salt": salt,
                               "password_hash": digest, "active": True}
            changed = True
    if changed or not USERS_FILE.exists():
        _save_users(users)
    return users


def _save_users(users):
    tmp = USERS_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(users, indent=2), encoding="utf-8")
    tmp.replace(USERS_FILE)


def _password_ok(password, record):
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), record["salt"].encode(), 180000).hex()
    return hmac.compare_digest(digest, record["password_hash"])



def user_perms(role: str, user: Optional[dict] = None) -> dict:
    """Effective permissions for a login. Admin always has everything."""
    if role == "admin":
        return {k: True for k in PERMISSION_KEYS}
    perms = dict(ROLE_DEFAULT_PERMS.get(role, ROLE_DEFAULT_PERMS["guest"]))
    saved = (user or {}).get("permissions")
    if isinstance(saved, dict):
        for k in PERMISSION_KEYS:
            if k in saved:
                perms[k] = bool(saved[k])
    if role != "accountant":              # guests have no company of their own -> never upload / edit
        perms["upload"] = False
        perms["edit"] = False
    if not perms["view"]:                 # without "view" nothing else makes sense
        perms.update({"upload": False, "download": False, "edit": False})
    return perms


def _perms_from_form(role, view, upload, download, edit) -> dict:
    base = ROLE_DEFAULT_PERMS[role]
    given = {"view": view, "upload": upload, "download": download, "edit": edit}
    return user_perms(role, {"permissions": {k: (base[k] if given[k] is None else given[k])
                                             for k in PERMISSION_KEYS}})


def _sign(payload: str) -> str:
    return hmac.new(SECRET, payload.encode(), hashlib.sha256).hexdigest()


def make_session(role: str, who: str = "") -> str:
    """who = accountant username (empty for admin / guest)."""
    payload = f"{role}|{who}|{int(time.time()) + SESSION_SECONDS}"
    return base64.urlsafe_b64encode(f"{payload}|{_sign(payload)}".encode()).decode()


def read_session(request: Request) -> Optional[tuple]:
    """(role, who) for a valid signed cookie, else None. role: admin / accountant / guest."""
    raw = request.cookies.get(SESSION_COOKIE)
    if not raw:
        return None
    try:
        role, who, exp, sig = base64.urlsafe_b64decode(raw.encode()).decode().split("|")
        if not hmac.compare_digest(sig, _sign(f"{role}|{who}|{exp}")):
            return None
        if int(exp) < time.time() or role not in ("admin", "guest", "accountant", "managed"):
            return None
        if role in ("accountant", "managed"):
            user = _load_users().get(who)
            if not user or not user.get("active", True):
                return None
            if role == "managed":
                return user.get("role", "guest"), who
            # Legacy sessions are accepted only while their managed record exists.
        return role, who
    except Exception:
        return None


async def auth_guard(request: Request, call_next):
    path = request.url.path
    # QuickBooks Web Connector authenticates with its own application credentials.
    if path == "/qbwc":
        return await call_next(request)
    if path in PUBLIC_PATHS:
        return await call_next(request)
    session = read_session(request)
    if session is None:
        if request.method == "GET" and not path.startswith("/api/"):
            return RedirectResponse("/login", status_code=303)
        return JSONResponse({"detail": "Please log in."}, status_code=401)
    role, who = session
    user = _load_users().get(who) if role in ("accountant", "managed") else None
    perms = user_perms(role, user)
    request.state.perms = perms

    def _match(method, prefix):
        return request.method == method and (path == prefix or (prefix.endswith("/") and path.startswith(prefix)))

    def _deny(msg):
        return JSONResponse({"detail": msg}, status_code=403)

    for method, prefix in ADMIN_ONLY:
        if _match(method, prefix) and role != "admin":
            return _deny("This login has view-only access.")
    if not perms["view"] and path not in ("/", "/api/me"):
        return _deny("You do not have permission to view sheets.")
    if _match("POST", "/api/sheets") and not perms["upload"]:
        return _deny("You do not have permission to upload sheets.")
    if _match("PUT", "/api/sheets/") and not (perms["upload"] or perms["edit"]):
        return _deny("You do not have permission to edit sheets.")
    if _match("DELETE", "/api/sheets/") and not perms["edit"]:
        return _deny("You do not have permission to edit or delete sheets.")
    if request.method == "POST" and path in DOWNLOAD_PATHS and not perms["download"]:
        return _deny("You do not have permission to download reports.")
    request.state.role = role
    request.state.scope = (user.get("company") if role == "accountant" else None) if user else None
    return await call_next(request)


def _login_html(error: str = "", tab: str = "admin") -> str:
    guest_pw = ('<label>Guest password</label><input name="password" type="password" required>'
                if GUEST_PASSWORD else "")
    return (LOGIN_PAGE
            .replace("__ERROR__", f'<div class="err">{error}</div>' if error else "")
            .replace("__TAB_ADMIN__", "active" if tab == "admin" else "")
            .replace("__TAB_ACCT__", "active" if tab == "accountant" else "")
            .replace("__TAB_GUEST__", "active" if tab == "guest" else "")
            .replace("__GUEST_PW__", guest_pw))


@router.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    if read_session(request):
        return RedirectResponse("/", status_code=303)
    return _login_html()


@router.post("/login")
async def login_submit(role: str = Form(...), username: str = Form(""), password: str = Form("")):
    who = ""
    if role == "admin":
        ok = hmac.compare_digest(username, ADMIN_USERNAME) and hmac.compare_digest(password, ADMIN_PASSWORD)
        if not ok:
            return HTMLResponse(_login_html("Wrong administrator username or password.", "admin"), status_code=401)
    elif role == "accountant":
        who = username.strip()
        acct = _load_users().get(who)
        if not acct or not acct.get("active", True) or not _password_ok(password, acct):
            return HTMLResponse(_login_html("Wrong username or password, or this account is disabled.", "accountant"), status_code=401)
        role = "managed"
    elif role == "guest":
        if GUEST_PASSWORD and not hmac.compare_digest(password, GUEST_PASSWORD):
            return HTMLResponse(_login_html("Wrong guest password.", "guest"), status_code=401)
    else:
        raise HTTPException(status_code=400, detail="Invalid login type.")
    resp = RedirectResponse("/", status_code=303)
    resp.set_cookie(SESSION_COOKIE, make_session(role, who if role == "managed" else ""), max_age=SESSION_SECONDS, httponly=True, samesite="lax")
    return resp


@router.get("/logout")
def logout():
    resp = RedirectResponse("/login", status_code=303)
    resp.delete_cookie(SESSION_COOKIE)
    return resp


@router.get("/api/me")
def api_me(request: Request):
    return {"role": request.state.role, "permissions": request.state.perms}

# ---- Administrator: User Management ---------------------------------------
def _require_admin(request: Request):
    if getattr(request.state, "role", None) != "admin":
        raise HTTPException(status_code=403, detail="Administrator access required.")


@router.get("/api/admin/users")
def admin_list_users(request: Request):
    _require_admin(request)
    users = _load_users()
    return {"users": [{"username": name, "role": u.get("role", "guest"),
                       "company": u.get("company", ""), "active": u.get("active", True),
                       "permissions": user_perms(u.get("role", "guest"), u)}
                      for name, u in sorted(users.items())]}


@router.get("/api/admin/companies")
def admin_user_companies(request: Request):
    _require_admin(request)
    with META_LOCK:
        companies = load_companies()
    return {"companies": [{"id": c["id"], "name": c["name"], "enabled": c.get("enabled", True)} for c in companies]}


@router.post("/api/admin/users")
def admin_create_user(request: Request, username: str = Form(...), password: str = Form(...),
                      role: str = Form(...), company: str = Form(""),
                      perm_view: Optional[bool] = Form(None), perm_upload: Optional[bool] = Form(None),
                      perm_download: Optional[bool] = Form(None), perm_edit: Optional[bool] = Form(None)):
    _require_admin(request)
    username = username.strip()
    if not username or len(username) > 50 or not all(ch.isalnum() or ch in "._-" for ch in username):
        raise HTTPException(400, "Username must be 1–50 characters (letters, numbers, dot, _ or -).")
    if len(password) < 8:
        raise HTTPException(400, "Password must contain at least 8 characters.")
    if role not in ("accountant", "guest"):
        raise HTTPException(400, "Choose Accountant or Guest access.")
    if role == "accountant":
        with META_LOCK:
            companies = load_companies()
        if not any(c["name"] == company and c.get("enabled", True) for c in companies):
            raise HTTPException(400, "Select an enabled company for this accountant.")
    else:
        company = ""
    with META_LOCK:
        users = _load_users()
        if username in users:
            raise HTTPException(409, "That username already exists.")
        salt = secrets.token_hex(16)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 180000).hex()
        users[username] = {"role": role, "company": company, "salt": salt,
                           "password_hash": digest, "active": True,
                           "permissions": _perms_from_form(role, perm_view, perm_upload, perm_download, perm_edit)}
        _save_users(users)
    return {"ok": True}


@router.put("/api/admin/users/{username}")
def admin_update_user(username: str, request: Request, role: str = Form(...),
                      company: str = Form(""), password: str = Form(""), active: bool = Form(True),
                      perm_view: Optional[bool] = Form(None), perm_upload: Optional[bool] = Form(None),
                      perm_download: Optional[bool] = Form(None), perm_edit: Optional[bool] = Form(None)):
    _require_admin(request)
    if role not in ("accountant", "guest"):
        raise HTTPException(400, "Choose Accountant or Guest access.")
    if password and len(password) < 8:
        raise HTTPException(400, "New password must contain at least 8 characters.")
    if role == "accountant":
        with META_LOCK:
            companies = load_companies()
        if not any(c["name"] == company and c.get("enabled", True) for c in companies):
            raise HTTPException(400, "Select an enabled company for this accountant.")
    else:
        company = ""
    with META_LOCK:
        users = _load_users()
        user = users.get(username)
        if not user:
            raise HTTPException(404, "User not found.")
        user.update({"role": role, "company": company, "active": active,
                     "permissions": _perms_from_form(role, perm_view, perm_upload, perm_download, perm_edit)})
        if password:
            salt = secrets.token_hex(16)
            user.update({"salt": salt, "password_hash": hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 180000).hex()})
        _save_users(users)
    return {"ok": True}


@router.delete("/api/admin/users/{username}")
def admin_revoke_user(username: str, request: Request):
    _require_admin(request)
    with META_LOCK:
        users = _load_users()
        if username not in users:
            raise HTTPException(404, "User not found.")
        del users[username]
        _save_users(users)
    return {"ok": True}
