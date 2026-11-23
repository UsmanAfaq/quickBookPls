"""Settings, folders and the signing secret. Environment variables (all optional):
ADMIN_USERNAME, ADMIN_PASSWORD, GUEST_PASSWORD, OASIS_PASSWORD, KRAKEN_PASSWORD, SECRET_KEY, PLS_DATA_DIR.
"""
import os
import secrets
import threading
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("PLS_DATA_DIR", str(BASE_DIR / "data")))
SHEETS_DIR = DATA_DIR / "sheets"
META_FILE = DATA_DIR / "sheets.json"
COMPANIES_FILE = DATA_DIR / "companies.json"
SHEETS_DIR.mkdir(parents=True, exist_ok=True)
META_LOCK = threading.Lock()

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")
GUEST_PASSWORD = os.environ.get("GUEST_PASSWORD", "")

# ---------------------------------------------------------------------------
# Accountant logins (view-only, each sees ONE company only).
#
# Username  : oasis   / kraken    (change with OASIS_USERNAME / KRAKEN_USERNAME)
# Password  : the value below, unless OASIS_PASSWORD / KRAKEN_PASSWORD is set
#             in the environment (the environment value always wins).
#
# >>> CHANGE THESE TWO PASSWORDS before using the tool. Use English letters,
# >>> numbers and symbols only. Do not share this file / push it to GitHub
# >>> while real passwords are written here.
# ---------------------------------------------------------------------------
DEFAULT_ACCOUNTANT_PASSWORDS = {
    "OASIS": "Oasis@2026",
    "KRAKEN": "Kraken@2026",
}


def _accountants() -> dict:
    """Studio accountants: each can see ONE company only (view-only).
    Company names must match the company name in the dashboard."""
    out = {}
    for prefix, company in (("OASIS", "Oasis Studio"), ("KRAKEN", "Kraken Studios")):
        password = os.environ.get(f"{prefix}_PASSWORD") or DEFAULT_ACCOUNTANT_PASSWORDS.get(prefix, "")
        if password:
            username = os.environ.get(f"{prefix}_USERNAME", prefix.lower())
            out[username] = {"password": password,
                             "company": os.environ.get(f"{prefix}_COMPANY", company)}
    return out


ACCOUNTANTS = _accountants()
SESSION_COOKIE = "pls_session"
SESSION_SECONDS = 12 * 3600


def _secret_key() -> bytes:
    env = os.environ.get("SECRET_KEY")
    if env:
        return env.encode()
    f = DATA_DIR / ".secret_key"
    if f.exists():
        return f.read_bytes()
    key = secrets.token_bytes(32)
    f.write_bytes(key)
    return key


SECRET = _secret_key()