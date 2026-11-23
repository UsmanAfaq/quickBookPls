"""sheets.json (metadata of uploaded sheets) and companies.json (workspaces) read / write."""
import json
import time
import uuid

from .config import COMPANIES_FILE, META_FILE

DEFAULT_COMPANY_NAME = "Ozi Group"   # existing sheets (uploaded before companies existed) go here


def load_meta() -> list:
    if not META_FILE.exists():
        return []
    try:
        return json.loads(META_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []


def save_meta(meta: list) -> None:
    tmp = META_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    tmp.replace(META_FILE)


def _read_companies() -> list:
    if not COMPANIES_FILE.exists():
        return []
    try:
        return json.loads(COMPANIES_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []


def save_companies(companies: list) -> None:
    tmp = COMPANIES_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(companies, indent=2), encoding="utf-8")
    tmp.replace(COMPANIES_FILE)


def new_company(name: str, enabled: bool = True) -> dict:
    return {"id": uuid.uuid4().hex[:12], "name": name.strip(), "enabled": enabled,
            "created_at": int(time.time())}


def load_companies() -> list:
    """Companies list. First run: creates the default company and moves every
    existing sheet (which has no company_id yet) into it. Caller holds META_LOCK."""
    companies = _read_companies()
    meta = load_meta()
    orphans = [m for m in meta if not m.get("company_id")]
    if not companies and (orphans or not COMPANIES_FILE.exists()):
        companies = [new_company(DEFAULT_COMPANY_NAME)]
        save_companies(companies)
    if orphans and companies:
        default_id = companies[0]["id"]
        for m in orphans:
            m["company_id"] = default_id
        save_meta(meta)
    # Sheets that point at a company id missing from companies.json (e.g. sheets.json
    # copied without companies.json) would be invisible -> re-create that company.
    known = {c["id"] for c in companies}
    missing = []
    for m in meta:
        cid = m.get("company_id")
        if cid and cid not in known and cid not in missing:
            missing.append(cid)
    if missing:
        for n, cid in enumerate(missing, 1):
            c = new_company(DEFAULT_COMPANY_NAME if n == 1 and not companies else f"Recovered company {n}")
            c["id"] = cid
            companies.append(c)
        save_companies(companies)
    return companies