from flask import Flask, jsonify, render_template, Response, request, send_from_directory, session, redirect, url_for, g
import pandas as pd
import re
from datetime import datetime
import os, time, queue, threading, uuid, hashlib, secrets, hmac
from datetime import timedelta
from urllib.parse import urlparse
from werkzeug.utils import secure_filename

from watchdog.observers import Observer
from watchdog.events    import FileSystemEventHandler

from students_data import (
    load_students,
    find_students_for_university,
    canonicalize_exchange_university,
    INTERNSHIP_ACTIVITIES,
)
from incoming_students import load_incoming_students
import auth as auth_store

def _load_or_create_secret_key():
    key_path = os.path.join(os.path.dirname(__file__), "data", "secret_key.txt")
    if os.environ.get("OIA_HUB_SECRET_KEY"):
        return os.environ["OIA_HUB_SECRET_KEY"]
    if os.path.isfile(key_path):
        with open(key_path, "r", encoding="utf-8") as f:
            existing = f.read().strip()
            if existing:
                return existing
    key = secrets.token_hex(32)
    os.makedirs(os.path.dirname(key_path), exist_ok=True)
    with open(key_path, "w", encoding="utf-8") as f:
        f.write(key)
    return key


app = Flask(__name__)
app.secret_key = _load_or_create_secret_key()

# Behind a reverse proxy (nginx / IIS / Apache on a university server), trust
# its X-Forwarded-* headers so login throttling sees each user's real IP and
# redirects use the public https:// address. Only enable when a proxy is in front.
if os.environ.get("OIA_HUB_BEHIND_PROXY") == "1":
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",          # cookie isn't sent on cross-site POSTs
    SESSION_COOKIE_SECURE=os.environ.get("OIA_HUB_HTTPS") == "1",  # set when served over HTTPS
    PERMANENT_SESSION_LIFETIME=timedelta(hours=int(os.environ.get("OIA_HUB_SESSION_HOURS", "16"))),
    # The cookie's expiry is fixed at sign-in rather than pushed back on every
    # click, so a sign-in lasts exactly SESSION_HOURS and then asks again.
    SESSION_REFRESH_EACH_REQUEST=False,
)
SESSION_MAX_SECONDS = int(app.config["PERMANENT_SESSION_LIFETIME"].total_seconds())

STUDENTS_EXCEL = os.path.join(os.path.dirname(__file__), "data", "Student_List_Abroad.xlsx")
STUDENTS = load_students(STUDENTS_EXCEL)
STUDENTS_BY_ID = {s["id"]: s for s in STUDENTS}

INCOMING_STUDENTS = load_incoming_students()
INCOMING_STUDENTS_BY_ID = {s["id"]: s for s in INCOMING_STUDENTS}

# Combined lookup so /api/students/<id> and the document-gallery routes work
# for both outbound (PU -> partner) and inbound (partner -> PU) students.
ALL_STUDENTS_BY_ID = {**STUDENTS_BY_ID, **INCOMING_STUDENTS_BY_ID}


def reload_roster_data():
    """Re-reads the outbound/inbound student rosters from disk.

    load_universities()/load_masters() re-read their Excel files on every
    call already, so a fresh upload of those shows up on the very next
    request. STUDENTS/INCOMING_STUDENTS are computed once at import time
    instead (they're merged from several workbooks), so this must be called
    explicitly after an upload replaces Student_List_Abroad.xlsx,
    Semester_Exchange_Consolidated_Data.xlsx, Incoming_Student_Data.xlsx, or
    Internship.xlsx for the new data to appear without restarting the app.
    """
    global STUDENTS, STUDENTS_BY_ID, INCOMING_STUDENTS, INCOMING_STUDENTS_BY_ID, ALL_STUDENTS_BY_ID
    STUDENTS = load_students(STUDENTS_EXCEL)
    STUDENTS_BY_ID = {s["id"]: s for s in STUDENTS}
    INCOMING_STUDENTS = load_incoming_students()
    INCOMING_STUDENTS_BY_ID = {s["id"]: s for s in INCOMING_STUDENTS}
    ALL_STUDENTS_BY_ID = {**STUDENTS_BY_ID, **INCOMING_STUDENTS_BY_ID}

DOCS_ROOT = os.path.join(os.path.dirname(__file__), "student-docs")
os.makedirs(DOCS_ROOT, exist_ok=True)
ALLOWED_DOC_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "webp", "pdf"}

EXCHANGE_DOCS_ROOT = os.path.join(os.path.dirname(__file__), "exchange-docs")
os.makedirs(EXCHANGE_DOCS_ROOT, exist_ok=True)

INTERNSHIP_DOCS_ROOT = os.path.join(os.path.dirname(__file__), "internship-docs")
os.makedirs(INTERNSHIP_DOCS_ROOT, exist_ok=True)


def _doc_dir(root, entity_id):
    path = os.path.join(root, secure_filename(entity_id))
    os.makedirs(path, exist_ok=True)
    return path


def _list_docs(root, entity_id, url_prefix):
    folder = _doc_dir(root, entity_id)
    docs = []
    for fname in sorted(os.listdir(folder)):
        parts = fname.split("__", 2)
        doc_type = parts[0].replace("_", " ") if len(parts) == 3 else "Document"
        original = parts[2] if len(parts) == 3 else fname
        docs.append({
            "filename": fname,
            "doc_type": doc_type,
            "original_name": original,
            "url": f"/{url_prefix}/{secure_filename(entity_id)}/{fname}",
        })
    return docs


def _allowed_doc(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_DOC_EXTENSIONS


def _student_doc_dir(student_id):
    path = os.path.join(DOCS_ROOT, secure_filename(student_id))
    os.makedirs(path, exist_ok=True)
    return path


def _list_student_docs(student_id):
    folder = _student_doc_dir(student_id)
    docs = []
    for fname in sorted(os.listdir(folder)):
        # filenames are stored as "<DocType>__<uuid>__<original>.ext"
        parts = fname.split("__", 2)
        doc_type = parts[0].replace("_", " ") if len(parts) == 3 else "Document"
        original = parts[2] if len(parts) == 3 else fname
        docs.append({
            "filename": fname,
            "doc_type": doc_type,
            "original_name": original,
            "url": f"/student-docs/{secure_filename(student_id)}/{fname}",
        })
    return docs


EXCEL_FILE = os.path.join(os.path.dirname(__file__), "MoU List (1) (2).xlsx")
EXCEL_SHEET = "MoU List (2)"  # superset of "MoU List" — includes 5 new MoUs + 6 renewals
RANKING_EXCEL_FILE = os.path.join(
    os.path.dirname(__file__), "Masters_Program_University_data-17-Feb-2026-cleaned (1).xlsx"
)
MASTERS_EXCEL_CANDIDATES = [
    os.path.join(os.path.dirname(__file__), "Masters_Program_List.xlsx"),
    os.path.join(os.path.dirname(__file__), "masters_program_list.xlsx"),
]


def _masters_excel_path():
    for path in MASTERS_EXCEL_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None

# ── SSE broker ────────────────────────────────────────────────────────────────
_clients      = []          # list of queue.Queue, one per open browser tab
_clients_lock = threading.Lock()
_data_version = 0           # bumped on every detected Excel change


def _broadcast(event: str, data: str):
    msg = f"event: {event}\ndata: {data}\n\n"
    with _clients_lock:
        dead = []
        for q in _clients:
            try:
                q.put_nowait(msg)
            except queue.Full:
                dead.append(q)
        for q in dead:
            _clients.remove(q)


# ── Watchdog ──────────────────────────────────────────────────────────────────
class ExcelWatcher(FileSystemEventHandler):
    """
    Fires on any file-system event that touches our Excel file.
    Debounced to 1 s so multi-write saves (Excel, LibreOffice) don't spam.
    """
    def __init__(self):
        self._last = 0

    def _maybe_fire(self, path):
        watched = {os.path.abspath(EXCEL_FILE)}
        masters_path = _masters_excel_path()
        if masters_path:
            watched.add(os.path.abspath(masters_path))
        if os.path.abspath(path) not in watched:
            return
        now = time.time()
        if now - self._last < 1.0:
            return
        self._last = now
        global _data_version
        _data_version += 1
        ver = _data_version
        time.sleep(0.6)
        _broadcast("update", str(ver))
        print(f"[watcher] Excel saved  →  version {ver}", flush=True)

    def on_modified(self, e): self._maybe_fire(getattr(e, "src_path", ""))
    def on_created (self, e): self._maybe_fire(getattr(e, "src_path", ""))  # rename-save


def start_watcher():
    handler  = ExcelWatcher()
    observer = Observer()
    watch_dir = os.path.dirname(os.path.abspath(EXCEL_FILE)) or "."
    observer.schedule(handler, watch_dir, recursive=False)
    observer.daemon = True
    observer.start()
    print(f"[watcher] Watching  →  {watch_dir}", flush=True)

def _status(row) -> str:
    raw = row.get("Valid_Upto")
    try:
        if pd.isna(raw):
            return "Unknown"
        days = (pd.to_datetime(raw) - datetime.now()).days
        if   days < 0:    return "Expired"
        elif days <= 365: return "Renewal Due"
        else:             return "Active"
    except Exception:
        return "Unknown"


def _date(raw, fmt: str) -> str:
    try:
        return pd.to_datetime(raw).strftime(fmt) if raw and str(raw).strip() else "—"
    except Exception:
        return "—"


def _clean_url(raw: str) -> str:
    """
    Normalise a raw website cell value from the new MoU List Excel.
    Handles:
      - 'Web: http://...'  / 'Website: https://...'  (with possible non-breaking space)
      - Plain 'http://...' or 'https://...'
      - Bare domain 'www.example.com' or 'example.edu.tw'
      - Email-only cells (no usable URL) -> return ''
    """
    import re
    val = str(raw).strip()
    # Strip 'Web:' / 'Website:' prefixes (including non-breaking spaces)
    val = re.sub(r'^[Ww]eb(?:site)?[\s\xa0]*:[\s\xa0]*', '', val).strip()
    if not val or val.lower() == "nan":
        return ""
    # Discard cells that are only an e-mail address
    if re.match(r'^[^@\s]+@[^@\s]+$', val):
        return ""
    # Already has a scheme
    if re.match(r'^https?://', val, re.I):
        return val
    # Bare domain / path - prepend https://
    if "." in val and " " not in val:
        return "https://" + val.lstrip('/')
    return ""


def _website(row) -> str:
    for key in ("Website", "URL", "University Website", "Website URL", "Web"):
        raw = row.get(key, "")
        if not raw:
            continue
        url = _clean_url(str(raw))
        if url:
            return url
    return ""


def _short_name(name: str, row) -> str:
    explicit = str(row.get("Short_Name", row.get("Short Name", ""))).strip()
    if explicit:
        return explicit
    skip = {"of", "the", "and", "&", "for", "in", "at", "a", "an"}
    words = [
        w for w in name.replace("&", " and ").split()
        if w.lower() not in skip and w[0].isalnum()
    ]
    if len(words) >= 2:
        return "".join(w[0].upper() for w in words[:5])
    return name[:8].upper()


def _mou_number(uni_id: str) -> str:
    if not uni_id:
        return ""
    clean = str(uni_id).strip()
    if clean.endswith(".0"):
        clean = clean[:-2]
    return f"MOU-{clean}" if clean else ""


def _valid_year(valid_upto: str) -> str:
    try:
        return pd.to_datetime(valid_upto).strftime("%Y")
    except Exception:
        return valid_upto if valid_upto and valid_upto != "—" else ""


def _clean(val) -> str:
    """Return stripped string, treating '*' / '-' / invisible chars as empty."""
    import unicodedata
    s = str(val).strip() if val is not None else ""
    # Remove zero-width spaces and other invisible unicode
    s = "".join(c for c in s if unicodedata.category(c) != "Cf")
    s = s.strip()
    return "" if s in ("*", "-", "nan", "None") else s


def _norm_mou_key(raw) -> str:
    """Normalizes a 'MoU No.' cell to a bare-digit string for MATCHING only
    (e.g. '01' -> '1', 1.0 -> '1', '194' -> '194'). Keeps '01'-style display
    values ('MOU-01') untouched elsewhere — this key is only used to join
    contacts to their university row."""
    s = str(raw).strip()
    if not s or s in ("*", "nan", "None"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    try:
        return str(int(s))
    except ValueError:
        return s


def load_contacts() -> dict[str, dict]:
    """
    Returns a dict keyed by MoU number (as int string, e.g. '7')
    with keys: name, designation, email, number.
    Reads from the same MoU List Excel (header row = row 1).
    """
    try:
        # Read the same way load_universities() does (header=0, then promote
        # row 0 to column names) rather than pandas' own header=1. Reading
        # this column via header=1 lets pandas infer "MoU No." as float64
        # (1.0, 2.0, ...), which strips the leading zero the source uses
        # ("01", "02", ...) and made contacts fail to match ~half the MoUs.
        raw = pd.read_excel(EXCEL_FILE, sheet_name=EXCEL_SHEET, header=0)
        raw.columns = raw.iloc[0]
        df = raw.iloc[1:].reset_index(drop=True)
        df = df.fillna("")
        contacts: dict[str, dict] = {}
        for _, r in df.iterrows():
            mou_no = _norm_mou_key(r.get("MoU No.", ""))
            if not mou_no:
                continue
            contacts[mou_no] = {
                "contact_name":        _clean(r.get("Contact Person  Name and Designation", "")),
                "contact_designation": _clean(r.get("Designation", "")),
                "contact_email":       _clean(r.get("Contact Email ID", "")),
                "contact_number":      _clean(r.get("Contact Number", "")),
            }
        return contacts
    except Exception as exc:
        print(f"[contacts] Failed to load: {exc}", flush=True)
        return {}


EXTRA_CONTACTS_EXCEL = os.path.join(
    os.path.dirname(__file__), "data", "extra_sources", "Partner_and_Non_Partner_Contact_Data.xlsx"
)


def _norm_uni_key(name: str) -> str:
    return re.sub(r"\s+", " ", str(name or "")).strip().lower()


def load_extra_contact_numbers() -> dict[str, list[str]]:
    """
    Supplementary phone numbers keyed by normalized university name, pulled
    from the Partner / Non-Partner / Point-of-Contact tracker. Used to fill
    in a contact number whenever the main MoU List doesn't have one on file.
    Returns e.g. {"aalen university": ["+49 (0) 7361 576-1704 (Katrin Zinser)"]}
    """
    if not os.path.isfile(EXTRA_CONTACTS_EXCEL):
        return {}

    sheet_specs = [
        # (sheet name, header row index, university col, number col, name col)
        ("Sheet1",           1, "Name of the University", "Contact Number", None),
        ("Partner",          0, "University Name ",        "Mobile Number ", "Name"),
        ("Non Partner",      1, "University Name ",        "Phone Number ",  "Name"),
        ("Point-of-Contact", 0, "University Name ",        "Mobile Number ", "Name"),
    ]

    numbers: dict[str, list[str]] = {}
    try:
        xls = pd.ExcelFile(EXTRA_CONTACTS_EXCEL)
        for sheet, header_row, uni_col, num_col, name_col in sheet_specs:
            if sheet not in xls.sheet_names:
                continue
            usecols = None
            if sheet in ("Partner", "Point-of-Contact"):
                usecols = list(range(7))
            elif sheet == "Non Partner":
                usecols = list(range(7))
            df = pd.read_excel(xls, sheet_name=sheet, header=header_row, usecols=usecols)
            for _, r in df.iterrows():
                uni = _clean(r.get(uni_col, ""))
                num = _clean(r.get(num_col, "")) if num_col in df.columns else ""
                if not uni or not num:
                    continue
                key = _norm_uni_key(uni)
                person = _clean(r.get(name_col, "")) if name_col and name_col in df.columns else ""
                entry = f"{num} ({person})" if person else num
                bucket = numbers.setdefault(key, [])
                if entry not in bucket:
                    bucket.append(entry)
    except Exception as exc:
        print(f"[extra-contacts] Failed to load: {exc}", flush=True)
    return numbers


def load_universities() -> list[dict]:
    # Row 0 is the real header row
    raw = pd.read_excel(EXCEL_FILE, sheet_name=EXCEL_SHEET, header=0)
    raw.columns = raw.iloc[0]           # promote row-0 values to column names
    df = raw.iloc[1:].reset_index(drop=True)
    df = df.fillna("")

    COL_NAME       = "Name of the University"
    COL_SHORT      = "Short Name"
    COL_COUNTRY    = "Country"
    COL_PROGRAM    = "Program"
    COL_PURPOSE    = "Purpose"
    COL_ACTIVITIES = "Activities"
    COL_SIGNED     = "Signed Date"
    COL_VALID_UPTO = "Valid upto"
    COL_MOU_NO     = "MoU No."

    # Load contacts once keyed by MoU number, plus supplementary phone
    # numbers keyed by university name (from the Partner/Non-Partner tracker)
    contacts = load_contacts()
    extra_numbers = load_extra_contact_numbers()

    rows = []
    for _, r in df.iterrows():
        name = str(r.get(COL_NAME, "")).strip()
        if not name or name.lower() == "nan":
            continue

        program  = str(r.get(COL_PROGRAM,    "")).strip()
        purpose  = str(r.get(COL_PURPOSE,    "")).strip()
        activity = str(r.get(COL_ACTIVITIES, "")).strip()
        mou_no   = str(r.get(COL_MOU_NO,     "")).strip()
        if mou_no.endswith(".0"):
            mou_no = mou_no[:-2]

        # Build a compat dict so existing helpers (_status, _date, _short_name) still work
        compat = dict(r)
        compat["Valid_Upto"]  = r.get(COL_VALID_UPTO, "")
        compat["Signed_Date"] = r.get(COL_SIGNED, "")
        compat["Short_Name"]  = str(r.get(COL_SHORT, "")).strip()

        valid_upto  = _date(compat["Valid_Upto"],  "%d %b %Y")
        signed_date = _date(compat["Signed_Date"], "%d %b %Y")
        website     = _website(compat)

        # Attach contact from the same Excel row (same MoU No. key)
        contact = contacts.get(_norm_mou_key(mou_no), {})

        # Contact number: prefer the MoU List's own column; if it's blank
        # (or just a "*" placeholder), fall back to the Partner / Non-Partner
        # / Point-of-Contact tracker matched on university name.
        contact_number = contact.get("contact_number", "")
        if not contact_number:
            fallback = extra_numbers.get(_norm_uni_key(name), [])
            contact_number = "; ".join(fallback)

        rows.append({
            "name":                name,
            "country":             str(r.get(COL_COUNTRY, "")).strip(),
            "program":             program,
            "purpose":             purpose,
            "activity":            activity,
            "status":              _status(compat),
            "valid_upto":          valid_upto,
            "valid_year":          _valid_year(valid_upto),
            "signed_date":         signed_date,
            "uni_id":              mou_no,
            "mou_number":          f"MOU-{mou_no}" if mou_no else "",
            "short_name":          _short_name(name, compat),
            "website":             website,
            "primary_activity":    activity or program or purpose or "General MoU",
            "contact_name":        contact.get("contact_name", ""),
            "contact_designation": contact.get("contact_designation", ""),
            "contact_email":       contact.get("contact_email", ""),
            "contact_number":      contact_number,
        })
    return rows


def _cell(row, *keys, default="") -> str:
    for key in keys:
        if key in row and row[key] is not None and str(row[key]).strip():
            return str(row[key]).strip()
    return default


def _split_schools(raw: str) -> list[str]:
    if not raw:
        return []
    parts = []
    for chunk in str(raw).replace(";", ",").split(","):
        label = chunk.strip()
        if label:
            parts.append(label)
    return parts


def load_masters_from_excel(path: str) -> list[dict]:
    df = pd.read_excel(path, sheet_name=0)
    df = df.fillna("")
    grouped: dict[str, dict] = {}

    for _, r in df.iterrows():
        row = r.to_dict()
        name = _cell(row, "University", "Partner University", "Uni Name")
        if not name:
            continue
        country = _cell(row, "Country", "Destination Country")
        city = _cell(row, "City", "Location")
        key = f"{name}|{country}"
        if key not in grouped:
            grouped[key] = {
                "id": _cell(row, "Uni_ID", "University ID") or key,
                "name": name,
                "country": country,
                "city": city,
                "programs": [],
            }

        program = {
            "name": _cell(row, "Program", "Program Name", "Course"),
            "opportunity": _cell(row, "Opportunity", "Pathway", "Opportunity Type", default="Direct Masters"),
            "puSchools": _split_schools(_cell(row, "PU School", "PU_School", "PU Schools", "Eligible Schools")),
            "duration": _cell(row, "Duration", "Program Duration"),
            "scholarship": _cell(row, "Scholarship", "Scholarship / Discount", "Discount"),
            "ranking": _cell(row, "Ranking", "QS/THE Ranking", "Rank"),
            "courseLink": _cell(row, "Course Link", "Course_Link", "Program Link", "Link"),
        }
        if program["name"]:
            grouped[key]["programs"].append(program)

    return [u for u in grouped.values() if u["programs"]]


def load_masters_from_mou() -> list[dict]:
    """Build a masters-oriented view from MoU rows that mention masters pathways."""
    grouped: dict[str, dict] = {}
    for row in load_universities():
        blob = f"{row.get('activity', '')} {row.get('program', '')} {row.get('purpose', '')}".lower()
        if "master" not in blob:
            continue
        key = f"{row['name']}|{row['country']}"
        if key not in grouped:
            grouped[key] = {
                "id": row.get("uni_id") or key,
                "name": row["name"],
                "country": row["country"],
                "city": "",
                "programs": [],
            }
        grouped[key]["programs"].append({
            "name": row.get("program") or row.get("purpose") or "Masters pathway",
            "opportunity": "Direct Masters" if "direct" in blob else "Pathway / MoU-linked",
            "puSchools": [],
            "duration": "",
            "scholarship": "",
            "ranking": "",
            "courseLink": "",
        })
    return [u for u in grouped.values() if u["programs"]]


_RANKING_CACHE = {"mtime": None, "data": None}


def _normalize_uni_name(text: str) -> str:
    text = (text or "").lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def load_ranking_data() -> list[dict]:
    """Reads QS/THE ranking + program data from the 'Complete Program list'
    sheet of the ranking-data workbook. Cached until the file changes."""
    if not os.path.isfile(RANKING_EXCEL_FILE):
        return []
    mtime = os.path.getmtime(RANKING_EXCEL_FILE)
    if _RANKING_CACHE["mtime"] == mtime and _RANKING_CACHE["data"] is not None:
        return _RANKING_CACHE["data"]

    df = pd.read_excel(RANKING_EXCEL_FILE, sheet_name="Complete Program list 17-Feb-26")
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    df = df.fillna("")

    rows = []
    for _, r in df.iterrows():
        row = r.to_dict()
        name = str(row.get("University Name", "")).strip()
        if not name:
            continue
        rows.append({
            "university": name,
            "norm": _normalize_uni_name(name),
            "country": str(row.get("Country", "")).strip(),
            "city": str(row.get("City", "")).strip(),
            "school": str(row.get("School at PU", "")).strip(),
            "program": str(row.get("Program offered at Partner University", "")).strip(),
            "opportunity": str(row.get("Opportunity", "")).strip(),
            "duration": str(row.get("Program Duration", "")).strip(),
            "scholarship": str(row.get("Partnership discount on Tuition fee", "")).strip(),
            "courseLink": str(row.get("Website link for course /fees/eligibility", "")).strip(),
            "universityWebsite": str(row.get("University Website", "")).strip(),
            "qsRanking": str(row.get("QS Ranking World Ranking", "")).strip(),
            "theRanking": str(row.get("THE Ranking World Ranking", "")).strip(),
        })

    _RANKING_CACHE["mtime"] = mtime
    _RANKING_CACHE["data"] = rows
    return rows


def _find_ranking_rows(ranking_data: list[dict], university_name: str) -> list[dict]:
    """Matches a university name against the ranking workbook's rows.

    NOTE: this used to also accept a "token subset" match (e.g. matching if
    every leftover word in the shorter name appeared in the longer one,
    after stripping generic words like "university"/"of"/"the"/"college"/
    "institute"/"school"). That silently merged unrelated universities that
    happen to share one distinctive word — e.g. "University of Bristol"
    matched "University of West of England Bristol" purely because both
    reduce to a token set containing "bristol". That cross-contaminated PU
    School and QS ranking data between the two unrelated institutions.
    Matching is now restricted to an exact normalized name, or one full
    normalized name being a contiguous substring of the other (which still
    catches real variants like a trailing campus qualifier) — no more
    leftover-word-set guessing.
    """
    target = _normalize_uni_name(university_name)
    if not target:
        return []

    matches = []
    for r in ranking_data:
        uni = r["norm"]
        if target == uni or target in uni or uni in target:
            matches.append(r)
    return matches


_PU_SCHOOL_BASE = {
    "soc": "School of Commerce",
    "socse": "School of Computer Science and Engineering",
    "sol": "School of Law",
    "sois": "School of Information Science",
    "sod": "School of Design",
    "som": "School of Management",
    "soms": "School of Media Studies",
    "soe": "School of Engineering",
}
_PU_SCHOOL_BRANCH = {
    "civil": "Civil", "civel": "Civil",
    "mech": "Mechanical",
    "ece": "ECE",
    "eee": "EEE",
    "pet": "Petroleum",
}


def _map_pu_school_code(raw: str) -> str:
    """Maps PU's internal school codes (e.g. 'SOE - Civil', 'SOCSE') to a
    readable full name, e.g. 'School of Engineering (Civil)'."""
    if not raw:
        return ""
    parts = re.split(r"[-–]", raw, maxsplit=1)
    base_key = re.sub(r"[^a-z]", "", parts[0].strip().lower())
    base = _PU_SCHOOL_BASE.get(base_key)
    if not base:
        return raw.strip()

    if len(parts) > 1:
        branch_key = re.sub(r"m\s*tech", "", parts[1].strip().lower())
        branch_key = re.sub(r"[^a-z]", "", branch_key)
        branch = _PU_SCHOOL_BRANCH.get(branch_key)
        if branch:
            return f"{base} ({branch})"
    return base


def _normalize_qs_ranking(raw: str) -> str:
    """For range values like '761-770' or '#100-150', keep the upper bound
    ('#770'). Single numeric values are kept as-is (with a '#' prefix).
    Non-numeric values (e.g. 'Not Ranked') are returned unchanged."""
    if not raw:
        return raw
    text = raw.strip()
    m = re.match(r"^#?\s*(\d+)\s*[-–]\s*(\d+)$", text)
    if m:
        return f"#{m.group(2)}"
    m = re.match(r"^#?\s*(\d+)$", text)
    if m:
        return f"#{m.group(1)}"
    return text


def enrich_with_rankings(universities: list[dict]) -> list[dict]:
    """Fills in QS ranking and PU school mapping on each university's
    programs by matching against the ranking workbook."""
    ranking_data = load_ranking_data()
    if not ranking_data:
        return universities

    for uni in universities:
        matches = _find_ranking_rows(ranking_data, uni["name"])
        if not matches:
            continue

        qs_values = [m["qsRanking"] for m in matches if m["qsRanking"]]
        best_qs = _normalize_qs_ranking(qs_values[0]) if qs_values else ""

        schools = []
        for m in matches:
            mapped = _map_pu_school_code(m.get("school", ""))
            if mapped and mapped not in schools:
                schools.append(mapped)

        for program in uni["programs"]:
            # Always replace with the freshly computed QS ranking from the
            # ranking workbook (not just when blank) — it's the authoritative
            # source, so a stale/blank value already on the program should
            # not win over it.
            if best_qs:
                program["ranking"] = best_qs
            if schools and not program.get("puSchools"):
                program["puSchools"] = list(schools)

    return universities


def load_masters_from_ranking() -> list[dict]:
    """Builds the Masters university/program list directly from the real
    'Complete Program list' sheet (the same workbook already parsed by
    load_ranking_data()). This is the actual per-program dataset OIA
    maintains — real program names, durations, scholarships, course links,
    and QS/THE rankings — so it's used as the primary source for
    /api/masters instead of the much thinner MoU-derived fallback."""
    ranking_data = load_ranking_data()
    if not ranking_data:
        return []

    grouped: dict[str, dict] = {}
    for r in ranking_data:
        name = r["university"]
        if not name:
            continue
        country = r["country"]
        key = f"{name}|{country}"
        if key not in grouped:
            grouped[key] = {
                "id": key,
                "name": name,
                "country": country,
                "city": r.get("city", ""),
                "website": "",
                "programs": [],
            }

        if not grouped[key]["website"]:
            grouped[key]["website"] = _clean_url(r.get("universityWebsite", ""))

        program_name = r.get("program", "")
        if not program_name:
            continue

        mapped_school = _map_pu_school_code(r.get("school", ""))
        qs = _normalize_qs_ranking(r.get("qsRanking", ""))

        grouped[key]["programs"].append({
            "name": program_name,
            "opportunity": r.get("opportunity", "") or "Direct Masters",
            "puSchools": [mapped_school] if mapped_school else [],
            "duration": r.get("duration", ""),
            "scholarship": r.get("scholarship", ""),
            "ranking": qs or r.get("theRanking", ""),
            "courseLink": r.get("courseLink", ""),
        })

    # A university-level "University Website" isn't filled in for every row
    # in the source sheet. When it's blank, fall back to the first program's
    # course link so the card still has somewhere useful to send the user,
    # rather than a dead "With Links" tile.
    for uni in grouped.values():
        if not uni["website"]:
            for program in uni["programs"]:
                link = _clean_url(program.get("courseLink", ""))
                if link:
                    uni["website"] = link
                    break

    return [u for u in grouped.values() if u["programs"]]


def load_masters() -> list[dict]:
    masters_path = _masters_excel_path()
    if masters_path:
        try:
            return enrich_with_rankings(load_masters_from_excel(masters_path))
        except Exception as exc:
            print(f"[masters] Failed to read {masters_path}: {exc}", flush=True)

    primary = load_masters_from_ranking()
    if primary:
        # Merge in any MoU-listed university that mentions "master" but has
        # no row of its own in the program workbook yet (e.g. Cardiff
        # University), so it doesn't disappear from the Masters view just
        # because its full program details haven't been added to that sheet.
        existing = {_normalize_uni_name(u["name"]) for u in primary}
        extras = enrich_with_rankings(load_masters_from_mou())
        for uni in extras:
            norm_name = _normalize_uni_name(uni["name"])
            if norm_name not in existing:
                primary.append(uni)
                existing.add(norm_name)
        return primary

    # Last-resort fallback if the program workbook itself can't be read at
    # all (e.g. missing file) — same behaviour as before this fix.
    return enrich_with_rankings(load_masters_from_mou())


# ── Data Quality Report ─────────────────────────────────────────────────────
# Live-computed audit of the loaded datasets: how much is actually usable
# right now, plus a record of the bugs already fixed (with their before/after
# counts) and the gaps that still need either a new source file or manual
# correction in the Excel by OIA. Everything under "metric" keys below is
# recomputed on every request from the current in-memory data — nothing here
# is a snapshot, so it stays accurate as the underlying Excel files change.

def _pct(part, total):
    return round(100 * part / total, 1) if total else 0.0


def _is_blank(val) -> bool:
    return not str(val or "").strip()


def _valid_email(value) -> bool:
    value = str(value or "").strip()
    return bool(re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", value))


def _counter_rows(counter, limit=10):
    total = sum(counter.values()) or 1
    return [
        {"label": str(label), "value": int(value), "pct": _pct(value, total)}
        for label, value in counter.most_common(limit)
    ]


def build_data_quality_report() -> dict:
    """Build the live, drill-down-ready data-quality dashboard payload."""
    unis = load_universities()
    masters = load_masters()
    masters_programs_list = [p for u in masters for p in u["programs"]]
    masters_programs = len(masters_programs_list)

    # -- MoU / partner-university contact coverage --------------------------
    n_unis = len(unis)

    # -- Outbound student roster --------------------------------------------
    n_students = len(STUDENTS)
    missing_email_s = sum(1 for s in STUDENTS if _is_blank(s.get("email")))
    missing_country_s = sum(1 for s in STUDENTS if _is_blank(s.get("country")))
    missing_uid_s = sum(1 for s in STUDENTS if _is_blank(s.get("uid")))
    missing_activity_s = sum(1 for s in STUDENTS if _is_blank(s.get("activity")))
    missing_university_s = sum(1 for s in STUDENTS if _is_blank(s.get("university")))
    bad_email_s = sum(1 for s in STUDENTS if not _is_blank(s.get("email")) and not _valid_email(s.get("email")))
    exchange_count = sum(1 for s in STUDENTS if (s.get("activity") or "").strip().lower() == "semester exchange")
    internship_students_list = [s for s in STUDENTS if (s.get("activity") or "").strip().lower() in INTERNSHIP_ACTIVITIES]
    internship_count = len(internship_students_list)

    # -- Inbound (incoming) exchange roster ----------------------------------
    n_incoming = len(INCOMING_STUDENTS)

    # -- Rich analytics ------------------------------------------------------
    activity_counter = __import__("collections").Counter(
        (s.get("activity") or "Not recorded").strip() or "Not recorded" for s in STUDENTS
    )
    status_counter = __import__("collections").Counter(
        (s.get("status") or "Not recorded").strip() or "Not recorded" for s in STUDENTS
    )
    country_counter = __import__("collections").Counter(
        (s.get("country") or "Not recorded").strip() or "Not recorded" for s in STUDENTS
    )
    batch_counter = __import__("collections").Counter(
        (s.get("batch") or "Not recorded").strip() or "Not recorded" for s in STUDENTS
    )
    course_counter = __import__("collections").Counter(
        (s.get("course") or "Not recorded").strip() or "Not recorded" for s in STUDENTS
    )

    core_fields = [
        ("Email", lambda s: _valid_email(s.get("email"))),
        ("Country", lambda s: not _is_blank(s.get("country"))),
        ("UID / Roll No.", lambda s: not _is_blank(s.get("uid"))),
        ("University", lambda s: not _is_blank(s.get("university"))),
        ("Activity", lambda s: not _is_blank(s.get("activity"))),
        ("Course", lambda s: not _is_blank(s.get("course"))),
        ("Start date", lambda s: not _is_blank(s.get("start_date"))),
        ("End date", lambda s: not _is_blank(s.get("end_date"))),
    ]
    completeness = []
    complete_records = 0
    for label, checker in core_fields:
        good = sum(1 for s in STUDENTS if checker(s))
        completeness.append({"label": label, "value": n_students - good, "good": good, "of": n_students, "pct": _pct(good, n_students)})
    complete_records = sum(1 for s in STUDENTS if all(checker(s) for _, checker in core_fields))
    quality_score = round(sum(item["pct"] for item in completeness) / len(completeness), 1) if completeness else 0

    # MoU health/status distribution.
    mou_status_counter = __import__("collections").Counter(
        (u.get("status") or "Unknown").strip() or "Unknown" for u in unis
    )
    mou_countries = __import__("collections").Counter(
        (u.get("country") or "Not recorded").strip() or "Not recorded" for u in unis
    )

    total_core_cells = n_students * len(core_fields)
    filled_core_cells = sum(item["good"] for item in completeness)

    return {
        "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "overview": {
            "mou_universities": n_unis, "outbound_students": n_students, "incoming_students": len(INCOMING_STUDENTS),
            "masters_universities": len(masters), "masters_programs": masters_programs,
            "exchange_students": exchange_count, "internship_students": internship_count,
        },
        "quality": {
            "score": quality_score,
            "complete_records": complete_records,
            "incomplete_records": n_students - complete_records,
            "filled_core_cells": filled_core_cells,
            "total_core_cells": total_core_cells,
            "invalid_emails": bad_email_s,
            "missing_fields": missing_email_s + missing_country_s + missing_uid_s + missing_university_s,
        },
        "analytics": {
            "activity": _counter_rows(activity_counter, 12),
            "status": _counter_rows(status_counter, 10),
            "countries": _counter_rows(country_counter, 12),
            "batches": _counter_rows(batch_counter, 15),
            "courses": _counter_rows(course_counter, 10),
            "mou_status": _counter_rows(mou_status_counter, 8),
            "mou_countries": _counter_rows(mou_countries, 10),
            "completeness": completeness,
        },
    }


@app.route("/reports")
def reports_page():
    return render_template("reports.html")


@app.route("/api/reports/data-quality")
def api_data_quality_report():
    return jsonify(build_data_quality_report())


# ── Authentication & per-page permissions ──────────────────────────────────
# Every page in the Hub can be *viewed* by anyone signed in. Editing a page's
# data (currently: uploading/removing documents on Student Exchange,
# Internships, and per-student records under Universities) is restricted to
# that page's controller(s) or an admin. The admin manages who controls what
# from the Admin page.

PUBLIC_ENDPOINTS = {"login", "static"}

# Login throttling: after MAX_FAILED_LOGINS bad attempts for the same
# username from the same address, further attempts are refused for LOCKOUT_SECONDS.
MAX_FAILED_LOGINS = 5
LOCKOUT_SECONDS = 10 * 60
_failed_logins = {}            # (ip, username) -> [timestamps of recent failures]
_failed_logins_lock = threading.Lock()


def _login_key(username):
    return (request.remote_addr or "?", (username or "").strip().lower())


def _lockout_remaining(key):
    now = time.time()
    with _failed_logins_lock:
        recent = [t for t in _failed_logins.get(key, []) if now - t < LOCKOUT_SECONDS]
        _failed_logins[key] = recent
        if len(recent) >= MAX_FAILED_LOGINS:
            return int(LOCKOUT_SECONDS - (now - recent[0])) + 1
    return 0


def _record_failed_login(key):
    with _failed_logins_lock:
        _failed_logins.setdefault(key, []).append(time.time())


def _clear_failed_logins(key):
    with _failed_logins_lock:
        _failed_logins.pop(key, None)


def _safe_next_url(target):
    """Only allow redirects back into this site (blocks open redirects like //evil.com)."""
    if not target:
        return url_for("index")
    parsed = urlparse(target)
    if parsed.scheme or parsed.netloc or not target.startswith("/") or target.startswith("//") or "\\" in target:
        return url_for("index")
    return target


def csrf_token():
    token = session.get("_csrf")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_csrf"] = token
    return token


CSRF_PROTECTED_ENDPOINTS = {
    "login", "admin_create_user", "admin_update_permissions", "admin_delete_user",
    "api_data_uploads_access",
}


def current_user():
    """The signed-in user record, or None. Cached for the duration of the request."""
    if "_cu" in g:
        return g._cu
    uid = session.get("user_id")
    g._cu = auth_store.find_by_id(uid) if uid is not None else None
    return g._cu


@app.before_request
def _require_login():
    if request.method == "POST" and request.endpoint in CSRF_PROTECTED_ENDPOINTS:
        sent = request.form.get("csrf_token", "")
        expected = session.get("_csrf", "")
        if not expected or not hmac.compare_digest(sent.encode("utf-8"), expected.encode("utf-8")):
            if request.endpoint == "login":
                return render_template("login.html", error="Your sign-in form expired. Please try again.", next_url=_safe_next_url(request.form.get("next"))), 400
            return jsonify({"error": "invalid or missing CSRF token — reload the page and try again"}), 400

    if request.endpoint in PUBLIC_ENDPOINTS or request.endpoint is None:
        return None

    # Look the user up on every request so a session for an account that was
    # deleted (or whose record vanished) stops working immediately.
    user = current_user()
    # Every sign-in lasts SESSION_MAX_SECONDS (16 h by default), checked on the
    # server too so an old cookie can't outlive it. Sessions from before this
    # rule have no login time and are asked to sign in once.
    expired = False
    if user is not None:
        login_at = session.get("login_at")
        if not isinstance(login_at, (int, float)) or time.time() - login_at > SESSION_MAX_SECONDS:
            expired, user = True, None
    if user is None:
        session.clear()
        g._cu = None
        if request.path.startswith("/api/"):
            return jsonify({"error": "authentication required", "expired": expired}), 401
        args = {"next": request.full_path.rstrip("?")}
        if expired:
            args["expired"] = "1"
        return redirect(url_for("login", **args))
    return None


@app.after_request
def _no_store_for_auth_pages(resp):
    if request.endpoint in ("login", "logout", "admin_page") or (request.endpoint or "").startswith("admin_"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


@app.context_processor
def _inject_user():
    return {
        "current_user": current_user(),
        "csrf_token": csrf_token,
        "session_ends_at": int(session["login_at"] + SESSION_MAX_SECONDS) if session.get("login_at") else None,
        "session_hours": SESSION_MAX_SECONDS // 3600,
        "PAGE_LABELS": auth_store.PAGE_LABELS,
        "PAGE_KEYS": auth_store.PAGE_KEYS,
        "UPLOAD_SLOTS": UPLOAD_SLOTS,
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        key = _login_key(username)
        wait = _lockout_remaining(key)
        if wait:
            minutes = max(1, round(wait / 60))
            error = f"Too many failed attempts. Try again in about {minutes} minute{'s' if minutes != 1 else ''}."
            return render_template("login.html", error=error, next_url=_safe_next_url(request.form.get("next"))), 429
        user = auth_store.verify_login(username, password)
        if user:
            _clear_failed_logins(key)
            session.clear()                 # new session on login (prevents session fixation)
            session["user_id"] = user["id"]
            session["login_at"] = time.time()
            # "Remember me": a persistent cookie survives closing the browser; otherwise
            # the cookie is dropped when the browser closes. The server-side
            # SESSION_MAX_SECONDS limit applies either way.
            session.permanent = request.form.get("remember") == "1"
            return redirect(_safe_next_url(request.form.get("next")))
        _record_failed_login(key)
        error = "Incorrect username or password."
        return render_template("login.html", error=error, next_url=_safe_next_url(request.form.get("next"))), 401
    if current_user():
        return redirect(_safe_next_url(request.args.get("next")))
    notice = None
    if request.args.get("expired"):
        notice = f"For security, each sign-in lasts {SESSION_MAX_SECONDS // 3600} hours. Please sign in again to continue."
    return render_template("login.html", error=error, notice=notice, next_url=request.args.get("next") and _safe_next_url(request.args.get("next")) or "")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def _admin_page_response(error=None, status=200):
    return render_template(
        "admin.html",
        users=auth_store.list_users(),
        error=error,
        default_password_in_use=auth_store.default_admin_password_in_use(),
        # Shown on the Admin page so the policy on screen always matches the
        # policy in code, rather than being hard-coded into the template.
        security={
            "min_password_length": auth_store.MIN_PASSWORD_LENGTH,
            "max_failed_logins": MAX_FAILED_LOGINS,
            "lockout_minutes": LOCKOUT_SECONDS // 60,
            "session_hours": SESSION_MAX_SECONDS // 3600,
            "https_only": app.config["SESSION_COOKIE_SECURE"],
        },
    ), status


def _require_admin():
    """Returns None when the caller is an admin, otherwise the response to send."""
    user = current_user()
    if not user:
        return redirect(url_for("login", next=request.path))
    if not user.get("is_admin"):
        if request.method == "GET":
            return render_template("forbidden.html", message="Admin access is required to view this page."), 403
        return jsonify({"error": "admin access required"}), 403
    return None


@app.route("/admin")
def admin_page():
    denied = _require_admin()
    if denied is not None:
        return denied
    return _admin_page_response()


@app.route("/admin/users", methods=["POST"])
def admin_create_user():
    denied = _require_admin()
    if denied is not None:
        return denied
    controls = request.form.getlist("controls")
    try:
        auth_store.create_user(
            username=request.form.get("username", ""),
            password=request.form.get("password", ""),
            is_admin=bool(request.form.get("is_admin")),
            controls=controls,
        )
    except ValueError as exc:
        return _admin_page_response(str(exc), 400)
    return redirect(url_for("admin_page"))


@app.route("/admin/users/<int:user_id>/permissions", methods=["POST"])
def admin_update_permissions(user_id):
    denied = _require_admin()
    if denied is not None:
        return denied
    me = current_user()
    controls = request.form.getlist("controls")
    is_admin = bool(request.form.get("is_admin"))
    if user_id == me["id"] and not is_admin:
        return _admin_page_response("You can't remove your own admin access. Ask another admin to do it.", 400)
    try:
        auth_store.update_permissions(user_id, is_admin, controls)
        new_password = request.form.get("new_password", "")
        if new_password:
            auth_store.update_password(user_id, new_password)
    except ValueError as exc:
        return _admin_page_response(str(exc), 400)
    return redirect(url_for("admin_page"))


@app.route("/admin/users/<int:user_id>/delete", methods=["POST"])
def admin_delete_user(user_id):
    denied = _require_admin()
    if denied is not None:
        return denied
    if user_id == current_user()["id"]:
        return _admin_page_response("You can't delete your own account while signed in as it.", 400)
    try:
        auth_store.delete_user(user_id)
    except ValueError as exc:
        return _admin_page_response(str(exc), 400)
    return redirect(url_for("admin_page"))


# ── Per-page data uploads ───────────────────────────────────────────────────
# Lets a page's controller (or an admin) replace that page's source Excel
# file directly from the browser, instead of a developer manually copying
# the file in. Saving over the existing path means load_universities() /
# load_masters() (which re-read from disk every call) show the new data on
# the very next request; the merged student rosters are recomputed
# explicitly via reload_roster_data() since those are cached at import time.
INTERNSHIP_XLSX_PATH = os.path.join(os.path.dirname(__file__), "data", "extra_sources", "Internship.xlsx")
SEMESTER_EXCHANGE_XLSX_PATH = os.path.join(os.path.dirname(__file__), "data", "extra_sources", "Semester_Exchange_Consolidated_Data.xlsx")
INCOMING_XLSX_PATH = os.path.join(os.path.dirname(__file__), "data", "extra_sources", "Incoming_Student_Data.xlsx")

UPLOAD_SLOTS = {
    "universities": [
        {
            "key": "mou_list", "label": "MoU List workbook", "path": EXCEL_FILE,
            "required_sheet": EXCEL_SHEET,
            "hint": f'Must contain a sheet named "{EXCEL_SHEET}".',
        },
        {
            "key": "contacts", "label": "Partner & Non-Partner Contact Data", "path": EXTRA_CONTACTS_EXCEL,
            "required_sheet": None,
            "hint": "Optional — fills in phone numbers the MoU List doesn't have.",
        },
    ],
    "masters": [
        {
            "key": "masters", "label": "Masters Program University workbook", "path": RANKING_EXCEL_FILE,
            "required_sheet": "Complete Program list 17-Feb-26",
            "hint": 'Must contain a "Complete Program list 17-Feb-26" sheet.',
        },
    ],
    "student_exchange": [
        {
            "key": "outbound_roster", "label": "Student List Abroad (outbound roster)", "path": STUDENTS_EXCEL,
            "required_sheet": "Final List (1-9)",
            "hint": 'Must contain a "Final List (1-9)" sheet (a "Top 300" sheet is used too, if present).',
        },
        {
            "key": "semester_exchange", "label": "Semester Exchange Consolidated Data", "path": SEMESTER_EXCHANGE_XLSX_PATH,
            "required_sheet": "Overall Data",
            "hint": 'Must contain an "Overall Data" sheet with a literal "Semester Exchange" activity column.',
        },
        {
            "key": "inbound_roster", "label": "Incoming Student Data (inbound roster)", "path": INCOMING_XLSX_PATH,
            "required_sheet": "Student Data",
            "hint": 'Must contain a "Student Data" sheet.',
        },
    ],
    "internships": [
        {
            "key": "internships", "label": "Internship workbook", "path": INTERNSHIP_XLSX_PATH,
            "required_sheet": None,
            "hint": "Batch-tab roster; internship/virtual-internship/summer-camp/summer-school rows are pulled in automatically.",
        },
    ],
}


def _upload_counts(page_key):
    """A short human-readable summary of what's now loaded, for the success message."""
    if page_key == "universities":
        unis = load_universities()
        return f"{len(unis)} universities mapped."
    if page_key == "masters":
        masters = load_masters()
        programs = sum(len(u["programs"]) for u in masters)
        return f"{len(masters)} universities, {programs} programs mapped."
    if page_key == "student_exchange":
        outbound = len([s for s in STUDENTS if (s.get("activity") or "").strip().lower() == "semester exchange"])
        return f"{outbound} outbound Semester Exchange students, {len(INCOMING_STUDENTS)} inbound students mapped."
    if page_key == "internships":
        n = len([s for s in STUDENTS if (s.get("activity") or "").strip().lower() in INTERNSHIP_ACTIVITIES])
        return f"{n} internship students mapped."
    return "Mapped."


# ── Upload safety: keep previous versions + an audit log ─────────────────────
BACKUP_ROOT = os.path.join(os.path.dirname(__file__), "data", "backups")
BACKUPS_TO_KEEP = 30          # per workbook; oldest are pruned


def _backup_before_upload(page_key, slot):
    """Copy the current workbook into data/backups/<page>/<slot>/ before an
    upload overwrites it, so a wrong upload can always be rolled back."""
    import shutil
    src = slot["path"]
    if not os.path.exists(src):
        return
    folder = os.path.join(BACKUP_ROOT, page_key, slot["key"])
    os.makedirs(folder, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    shutil.copy2(src, os.path.join(folder, f"{stamp}__{os.path.basename(src)}"))
    old = sorted(f for f in os.listdir(folder) if "__" in f)
    for f in old[:-BACKUPS_TO_KEEP]:
        try:
            os.remove(os.path.join(folder, f))
        except OSError:
            pass


def _log_upload(message):
    os.makedirs(BACKUP_ROOT, exist_ok=True)
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  {message}\n"
    with open(os.path.join(BACKUP_ROOT, "upload_log.txt"), "a", encoding="utf-8") as f:
        f.write(line)
    print(f"[upload] {message}", flush=True)


# ── Data Uploads page: one place to replace every source workbook ────────────
UPLOAD_PAGE_ICONS = {"universities": "🏛️", "masters": "🎓", "student_exchange": "✈️", "internships": "💼"}
_LOG_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})  (\S+) replaced (.+) \((.+)\) via (\S+)$")


def _last_uploads():
    """{(page_key, slot_label): {"by": user, "at": timestamp}} from the upload log."""
    out = {}
    path = os.path.join(BACKUP_ROOT, "upload_log.txt")
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8") as f:
        for line in f:
            m = _LOG_RE.match(line.rstrip("\n"))
            if m:
                at, by, label, _fname, page = m.groups()
                out[(page, label)] = {"by": by, "at": at}
    return out


@app.route("/data-uploads")
def data_uploads_page():
    return render_template("data_uploads.html")


@app.route("/api/data-uploads")
def api_data_uploads():
    """The workbooks this user may replace, with what's currently loaded."""
    user = current_user()
    is_admin = bool(user and user.get("is_admin"))
    last = _last_uploads()
    all_users = auth_store.list_users() if is_admin else []
    pages = []
    for page_key, slots in UPLOAD_SLOTS.items():
        if not auth_store.can_control(user, page_key):
            continue
        items = []
        for slot in slots:
            exists = os.path.exists(slot["path"])
            backup_dir = os.path.join(BACKUP_ROOT, page_key, slot["key"])
            items.append({
                "key": slot["key"],
                "label": slot["label"],
                "hint": slot["hint"],
                "file": {
                    "name": os.path.basename(slot["path"]),
                    "exists": exists,
                    "size": os.path.getsize(slot["path"]) if exists else 0,
                    "modified": datetime.fromtimestamp(os.path.getmtime(slot["path"])).strftime("%Y-%m-%d %H:%M") if exists else None,
                },
                "last_upload": last.get((page_key, slot["label"])),
                "backups": len([f for f in os.listdir(backup_dir) if "__" in f]) if os.path.isdir(backup_dir) else 0,
            })
        entry = {
            "key": page_key,
            "label": auth_store.PAGE_LABELS.get(page_key, page_key),
            "icon": UPLOAD_PAGE_ICONS.get(page_key, "📄"),
            "slots": items,
        }
        if is_admin:
            # Who may upload here: every admin, plus users granted this page.
            entry["uploaders"] = [
                {"id": u["id"], "username": u["username"], "is_admin": bool(u.get("is_admin"))}
                for u in all_users if auth_store.can_control(u, page_key)
            ]
        pages.append(entry)
    payload = {"pages": pages, "is_admin": is_admin}
    if is_admin:
        payload["users"] = [{"id": u["id"], "username": u["username"]} for u in all_users if not u.get("is_admin")]
    return jsonify(payload)


@app.route("/api/data-uploads/access", methods=["POST"])
def api_data_uploads_access():
    """Admin only: allow or stop one user uploading data for one page."""
    me = current_user()
    if not (me and me.get("is_admin")):
        return jsonify({"error": "Only admins can change who uploads data."}), 403
    page_key = request.form.get("page", "")
    if page_key not in UPLOAD_SLOTS:
        return jsonify({"error": "Unknown page."}), 400
    try:
        user_id = int(request.form.get("user_id", ""))
    except ValueError:
        return jsonify({"error": "Unknown user."}), 400
    allow = request.form.get("allow") == "1"
    target = next((u for u in auth_store.list_users() if u["id"] == user_id), None)
    if target is None:
        return jsonify({"error": "That user no longer exists."}), 404
    if target.get("is_admin"):
        return jsonify({"error": f"{target['username']} is an admin and can already upload everywhere."}), 400
    controls = [c for c in (target.get("controls") or []) if c != page_key]
    if allow:
        controls.append(page_key)
    auth_store.update_permissions(user_id, False, controls)
    label = auth_store.PAGE_LABELS.get(page_key, page_key)
    verb = "allowed" if allow else "removed"
    _log_upload(f"{me['username']} {verb} {target['username']} upload access for {label}")
    return jsonify({
        "ok": True,
        "message": (f"{target['username']} can now upload data for {label}." if allow
                    else f"{target['username']} can no longer upload data for {label}."),
    })


@app.route("/api/page-data/<page_key>/<slot_key>", methods=["POST"])
def upload_page_data(page_key, slot_key):
    user = current_user()
    if not auth_store.can_control(user, page_key):
        label = auth_store.PAGE_LABELS.get(page_key, page_key)
        return jsonify({"error": f"You don't have permission to update {label}. Ask an admin to grant you control of this page."}), 403

    slots = UPLOAD_SLOTS.get(page_key, [])
    slot = next((s for s in slots if s["key"] == slot_key), None)
    if not slot:
        return jsonify({"error": "unknown upload slot"}), 404

    file = request.files.get("file")
    if not file or file.filename == "":
        return jsonify({"error": "no file provided"}), 400
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": "please upload an .xlsx file"}), 400

    # Validate before overwriting the live file, so a bad upload can't wipe
    # out good data — read it into memory first, check it opens and (if the
    # slot specifies one) contains the sheet the loader actually reads.
    import openpyxl
    from io import BytesIO
    raw_bytes = file.read()
    try:
        wb = openpyxl.load_workbook(BytesIO(raw_bytes), read_only=True)
    except Exception:
        return jsonify({"error": "That file couldn't be read as an Excel workbook — it may be corrupted or not a real .xlsx file."}), 400

    required = slot.get("required_sheet")
    if required:
        needle = required.strip().lower()
        if not any(needle in name.strip().lower() for name in wb.sheetnames):
            return jsonify({
                "error": f'This workbook has no sheet matching "{required}" — found: {", ".join(wb.sheetnames)}. Nothing was changed.'
            }), 400
    wb.close()

    os.makedirs(os.path.dirname(slot["path"]), exist_ok=True)
    _backup_before_upload(page_key, slot)
    # Write to a temp file, then swap it in, so nobody ever reads a half-written workbook.
    tmp_path = slot["path"] + ".uploading"
    with open(tmp_path, "wb") as f:
        f.write(raw_bytes)
    os.replace(tmp_path, slot["path"])
    uploader = (current_user() or {}).get("username", "?")
    _log_upload(f"{uploader} replaced {slot['label']} ({os.path.basename(slot['path'])}) via {page_key}")

    if page_key in ("student_exchange", "internships"):
        reload_roster_data()

    global _data_version
    _data_version += 1
    _broadcast("update", str(_data_version))

    return jsonify({"ok": True, "message": f"{slot['label']} updated. {_upload_counts(page_key)}"})


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/masters-data")
def masters_data():
    return render_template("masters-data.html")


@app.route("/students/<uni_id>")
def students_page(uni_id):
    return render_template("students.html", uni_id=uni_id)


@app.route("/student/<student_id>")
def student_detail_page(student_id):
    return render_template("student_detail.html", student_id=student_id)


@app.route("/student-exchange")
def student_exchange_page():
    return render_template("student_exchange.html")


@app.route("/student-exchange/<uni_id>")
def student_exchange_detail_page(uni_id):
    return render_template("student_exchange_detail.html", uni_id=uni_id)


@app.route("/api/student-exchange")
def api_student_exchange():
    # Semester_Exchange_Consolidated_Data.xlsx (Batch 01..08) carries a real
    # "Semester Exchange" activity label per student, so outbound uses that
    # directly rather than the broader "everything non-internship" fallback.
    exchange_students = [s for s in STUDENTS if (s.get("activity") or "").strip().lower() == "semester exchange"]
    universities = sorted({s["university"] for s in exchange_students if s.get("university")})
    countries = sorted({s["country"] for s in exchange_students if s.get("country")})
    return jsonify({
        "students": exchange_students,
        "total": len(exchange_students),
        "universities": universities,
        "countries": countries,
    })


def _group_exchange_universities(exchange_students):
    """Groups Semester Exchange students by their real host university
    (canonicalized to merge typos/abbreviation variants), independent of the
    separate 61-university Masters Data partner list — semester exchange
    happens with a different, broader set of institutions."""
    groups = {}
    for s in exchange_students:
        canonical = canonicalize_exchange_university(s.get("university"))
        if not canonical:
            continue
        key = canonical.lower()
        if key not in groups:
            groups[key] = {
                "id": "ex_" + hashlib.md5(key.encode("utf-8")).hexdigest()[:10],
                "name": canonical,
                "country": s.get("country", ""),
                "students": [],
            }
        groups[key]["students"].append(s)
        mapped_country = _exchange_country_name(s.get("country"), canonical)
        if mapped_country and (
            not groups[key]["country"]
            or groups[key]["country"].strip().casefold() == "bristol"
        ):
            groups[key]["country"] = mapped_country

    results = []
    for g in groups.values():
        matches = g["students"]
        programs = sorted({(m.get("program") or m.get("course") or "").strip() for m in matches if (m.get("program") or m.get("course"))})

        school_counts = {}
        for m in matches:
            school = (m.get("pu_school") or "").strip()
            if school:
                school_counts[school] = school_counts.get(school, 0) + 1
        pu_schools = [{"name": k, "count": v} for k, v in sorted(school_counts.items(), key=lambda kv: -kv[1])]
        unmapped_count = sum(1 for m in matches if not (m.get("pu_school") or "").strip())

        results.append({
            "id": g["id"],
            "name": g["name"],
            "country": g["country"],
            "student_count": len(matches),
            "programs": programs,
            "students": matches,
            "pu_schools": pu_schools,
            "unmapped_school_count": unmapped_count,
        })

    results.sort(key=lambda u: -u["student_count"])
    return results


def _group_incoming_universities(incoming_students):
    """Groups Inbound Exchange students by their home (sending) university,
    already canonicalized in incoming_students.py."""
    groups = {}
    for s in incoming_students:
        name = s.get("university")
        if not name:
            continue
        key = name.lower()
        if key not in groups:
            groups[key] = {
                "id": "in_" + hashlib.md5(key.encode("utf-8")).hexdigest()[:10],
                "name": name,
                "country": s.get("country", ""),
                "students": [],
            }
        groups[key]["students"].append(s)
        if not groups[key]["country"] and s.get("country"):
            groups[key]["country"] = s["country"]

    results = []
    for g in groups.values():
        matches = g["students"]
        programs = sorted({(m.get("program") or "").strip() for m in matches if (m.get("program") or "").strip()})
        results.append({
            "id": g["id"],
            "name": g["name"],
            "country": g["country"],
            "student_count": len(matches),
            "programs": programs,
            "students": matches,
            "pu_schools": [],
            "unmapped_school_count": 0,
        })

    results.sort(key=lambda u: -u["student_count"])
    return results


def _exchange_university_key(name):
    """Stable identity key for combining inbound and outbound institution lists."""
    value = re.sub(r"[^a-z0-9]+", " ", (name or "").casefold())
    return " ".join(value.split())


def _exchange_country_name(name, university=""):
    """Normalize known source-data variants before counting/displaying countries."""
    country = (name or "").strip()
    university_key = _exchange_university_key(university)
    # One outbound source row incorrectly stores the city in the Country field.
    if university_key in {"uwe bristol", "university of the west of england bristol"} and country.casefold() == "bristol":
        return "UK"
    aliases = {
        "united kingdom": "UK",
        "great britain": "UK",
        "united arab emirates": "UAE",
        "republic of korea": "South Korea",
        "korea republic of": "South Korea",
    }
    return aliases.get(country.casefold(), country)


@app.route("/api/student-exchange/stats")
def api_student_exchange_stats():
    """Combined Student Exchange hero totals across both directions.

    Universities are distinct institutions across inbound + outbound (an
    institution appearing in both directions is counted once); students are
    records from both rosters; countries are distinct normalized countries.
    """
    outbound_students = [
        s for s in STUDENTS
        if (s.get("activity") or "").strip().lower() == "semester exchange"
    ]
    outbound_universities = _group_exchange_universities(outbound_students)
    inbound_universities = _group_incoming_universities(INCOMING_STUDENTS)

    university_keys = {
        _exchange_university_key(u.get("name"))
        for u in outbound_universities + inbound_universities
        if u.get("name")
    }
    country_keys = {
        _exchange_country_name(s.get("country"), s.get("university")).casefold()
        for s in outbound_students + INCOMING_STUDENTS
        if _exchange_country_name(s.get("country"), s.get("university"))
    }

    return jsonify({
        "universities": len(university_keys),
        "students": len(outbound_students) + len(INCOMING_STUDENTS),
        "countries": len(country_keys),
        "breakdown": {
            "outbound_universities": len(outbound_universities),
            "inbound_universities": len(inbound_universities),
            "outbound_students": len(outbound_students),
            "inbound_students": len(INCOMING_STUDENTS),
        },
    })


@app.route("/api/student-exchange/universities")
def api_student_exchange_universities():
    direction = request.args.get("direction", "outbound")
    if direction == "inbound":
        results = _group_incoming_universities(INCOMING_STUDENTS)
        return jsonify({"universities": results, "total": len(results)})

    exchange_students = [s for s in STUDENTS if (s.get("activity") or "").strip().lower() == "semester exchange"]
    results = _group_exchange_universities(exchange_students)
    return jsonify({"universities": results, "total": len(results)})


@app.route("/api/student-exchange/university/<uni_id>")
def api_student_exchange_university_detail(uni_id):
    if uni_id.startswith("in_"):
        results = _group_incoming_universities(INCOMING_STUDENTS)
    else:
        exchange_students = [s for s in STUDENTS if (s.get("activity") or "").strip().lower() == "semester exchange"]
        results = _group_exchange_universities(exchange_students)
    result = next((u for u in results if u["id"] == uni_id), None)
    if not result:
        return jsonify({"error": "not found"}), 404
    return jsonify({"university": result})


@app.route("/api/student-exchange/<uni_id>/documents", methods=["GET", "POST"])
def api_exchange_documents(uni_id):
    if request.method == "GET":
        return jsonify({"documents": _list_docs(EXCHANGE_DOCS_ROOT, uni_id, "exchange-docs")})

    if not auth_store.can_control(current_user(), "student_exchange"):
        return jsonify({"error": "You don't have permission to modify Student Exchange. Ask an admin to grant you control of this page."}), 403

    file = request.files.get("file")
    doc_type = (request.form.get("doc_type") or "Document").strip() or "Document"
    if not file or file.filename == "":
        return jsonify({"error": "no file provided"}), 400
    if not _allowed_doc(file.filename):
        return jsonify({"error": "unsupported file type"}), 400

    safe_type = secure_filename(doc_type) or "Document"
    original = secure_filename(file.filename)
    stored_name = f"{safe_type}__{uuid.uuid4().hex[:10]}__{original}"
    folder = _doc_dir(EXCHANGE_DOCS_ROOT, uni_id)
    file.save(os.path.join(folder, stored_name))

    return jsonify({"documents": _list_docs(EXCHANGE_DOCS_ROOT, uni_id, "exchange-docs")})


@app.route("/api/student-exchange/<uni_id>/documents/<filename>", methods=["DELETE"])
def api_exchange_document_delete(uni_id, filename):
    if not auth_store.can_control(current_user(), "student_exchange"):
        return jsonify({"error": "You don't have permission to modify Student Exchange. Ask an admin to grant you control of this page."}), 403
    folder = _doc_dir(EXCHANGE_DOCS_ROOT, uni_id)
    target = os.path.join(folder, secure_filename(filename))
    if os.path.isfile(target):
        os.remove(target)
    return jsonify({"documents": _list_docs(EXCHANGE_DOCS_ROOT, uni_id, "exchange-docs")})


@app.route("/exchange-docs/<uni_id>/<filename>")
def serve_exchange_doc(uni_id, filename):
    return send_from_directory(_doc_dir(EXCHANGE_DOCS_ROOT, uni_id), filename)


# ── Internships ─────────────────────────────────────────────────────────────
# INTERNSHIP_ACTIVITIES (imported from students_data) is also used there to
# pull in the internship-type rows from data/extra_sources/Internship.xlsx
# that are missing from the main roster — see EXTRA_WORKBOOKS in that file.
#
# NOTE: "virtual internship" (25 students) and "summer school" (9 students) were
# previously excluded from every Internships page/API even though they're
# clearly internship-type engagements in the source roster — the old set only
# had "internship"/"summer camp" (which doesn't even appear in the data) and
# missed the actual values used. "Winter Study Tour" (2 students) is left out
# since it isn't clearly an internship; flag it to OIA to confirm categorization.


@app.route("/internships")
def internships_page():
    return render_template("internships.html")


@app.route("/internships/<host_id>")
def internship_detail_page(host_id):
    return render_template("internship_detail.html", host_id=host_id)


@app.route("/api/internships")
def api_internships():
    internship_students = [s for s in STUDENTS if (s.get("activity") or "").strip().lower() in INTERNSHIP_ACTIVITIES]
    hosts = sorted({s["university"] for s in internship_students if s.get("university")})
    countries = sorted({s["country"] for s in internship_students if s.get("country")})
    return jsonify({
        "students": internship_students,
        "total": len(internship_students),
        "hosts": hosts,
        "countries": countries,
    })


def _group_internship_hosts(internship_students):
    """Groups Internship / Virtual Internship students by their host organization
    (company, lab, or university). Funding type and eligibility criteria
    aren't fields in the source roster yet, so they're surfaced as
    'Not recorded' until OIA adds that data."""
    groups = {}
    for s in internship_students:
        name = (s.get("university") or "").strip()
        if not name:
            continue
        key = name.lower()
        if key not in groups:
            groups[key] = {
                "id": "in_" + hashlib.md5(key.encode("utf-8")).hexdigest()[:10],
                "name": name,
                "country": s.get("country", ""),
                "students": [],
            }
        groups[key]["students"].append(s)
        if not groups[key]["country"] and s.get("country"):
            groups[key]["country"] = s["country"]

    results = []
    for g in groups.values():
        matches = g["students"]
        programs = sorted({(m.get("program") or m.get("course") or "").strip() for m in matches if (m.get("program") or m.get("course"))})
        activity_types = sorted({(m.get("activity") or "").strip() for m in matches if m.get("activity")})

        school_counts = {}
        for m in matches:
            school = (m.get("pu_school") or "").strip()
            if school:
                school_counts[school] = school_counts.get(school, 0) + 1
        pu_schools = [{"name": k, "count": v} for k, v in sorted(school_counts.items(), key=lambda kv: -kv[1])]
        unmapped_count = sum(1 for m in matches if not (m.get("pu_school") or "").strip())

        results.append({
            "id": g["id"],
            "name": g["name"],
            "country": g["country"],
            "student_count": len(matches),
            "programs": programs,
            "activity_types": activity_types,
            "students": matches,
            "pu_schools": pu_schools,
            "unmapped_school_count": unmapped_count,
            "funding_type": "",       # not tracked in the source roster yet
            "eligibility_criteria": "",  # not tracked in the source roster yet
            "link": "",               # not tracked in the source roster yet
        })

    results.sort(key=lambda h: -h["student_count"])
    return results


@app.route("/api/internships/hosts")
def api_internship_hosts():
    internship_students = [s for s in STUDENTS if (s.get("activity") or "").strip().lower() in INTERNSHIP_ACTIVITIES]
    results = _group_internship_hosts(internship_students)
    return jsonify({"hosts": results, "total": len(results)})


@app.route("/api/internships/host/<host_id>")
def api_internship_host_detail(host_id):
    internship_students = [s for s in STUDENTS if (s.get("activity") or "").strip().lower() in INTERNSHIP_ACTIVITIES]
    results = _group_internship_hosts(internship_students)
    result = next((h for h in results if h["id"] == host_id), None)
    if not result:
        return jsonify({"error": "not found"}), 404
    return jsonify({"host": result})


@app.route("/api/internships/<host_id>/documents", methods=["GET", "POST"])
def api_internship_documents(host_id):
    if request.method == "GET":
        return jsonify({"documents": _list_docs(INTERNSHIP_DOCS_ROOT, host_id, "internship-docs")})

    if not auth_store.can_control(current_user(), "internships"):
        return jsonify({"error": "You don't have permission to modify Internships. Ask an admin to grant you control of this page."}), 403

    file = request.files.get("file")
    doc_type = (request.form.get("doc_type") or "Gallery Photo").strip() or "Gallery Photo"
    if not file or file.filename == "":
        return jsonify({"error": "no file provided"}), 400
    if not _allowed_doc(file.filename):
        return jsonify({"error": "unsupported file type"}), 400

    safe_type = secure_filename(doc_type) or "Gallery_Photo"
    original = secure_filename(file.filename)
    stored_name = f"{safe_type}__{uuid.uuid4().hex[:10]}__{original}"
    folder = _doc_dir(INTERNSHIP_DOCS_ROOT, host_id)
    file.save(os.path.join(folder, stored_name))

    return jsonify({"documents": _list_docs(INTERNSHIP_DOCS_ROOT, host_id, "internship-docs")})


@app.route("/api/internships/<host_id>/documents/<filename>", methods=["DELETE"])
def api_internship_document_delete(host_id, filename):
    if not auth_store.can_control(current_user(), "internships"):
        return jsonify({"error": "You don't have permission to modify Internships. Ask an admin to grant you control of this page."}), 403
    folder = _doc_dir(INTERNSHIP_DOCS_ROOT, host_id)
    target = os.path.join(folder, secure_filename(filename))
    if os.path.isfile(target):
        os.remove(target)
    return jsonify({"documents": _list_docs(INTERNSHIP_DOCS_ROOT, host_id, "internship-docs")})


@app.route("/internship-docs/<host_id>/<filename>")
def serve_internship_doc(host_id, filename):
    return send_from_directory(_doc_dir(INTERNSHIP_DOCS_ROOT, host_id), filename)


@app.route("/api/students")
def api_students():
    university = request.args.get("university", "").strip()
    if university:
        matches = find_students_for_university(STUDENTS, university)
        return jsonify({"students": matches, "total": len(matches)})
    return jsonify({"students": STUDENTS, "total": len(STUDENTS)})


@app.route("/api/students/<student_id>")
def api_student_detail(student_id):
    student = ALL_STUDENTS_BY_ID.get(student_id)
    if not student:
        return jsonify({"error": "not found"}), 404
    return jsonify({"student": student, "documents": _list_student_docs(student_id)})


@app.route("/api/students/<student_id>/documents", methods=["GET", "POST"])
def api_student_documents(student_id):
    if student_id not in ALL_STUDENTS_BY_ID:
        return jsonify({"error": "not found"}), 404

    if request.method == "GET":
        return jsonify({"documents": _list_student_docs(student_id)})

    # Incoming (inbound) students are only reachable through Student Exchange;
    # outbound students are reachable through Universities, Student Exchange,
    # or Internships depending on their activity, so any one of those three
    # page controllers may attach documents to them.
    user = current_user()
    if student_id in INCOMING_STUDENTS_BY_ID:
        allowed = auth_store.can_control(user, "student_exchange")
    else:
        allowed = any(
            auth_store.can_control(user, p) for p in ("universities", "student_exchange", "internships")
        )
    if not allowed:
        return jsonify({"error": "You don't have permission to modify student records. Ask an admin to grant you control of the relevant page."}), 403

    file = request.files.get("file")
    doc_type = (request.form.get("doc_type") or "Document").strip() or "Document"
    if not file or file.filename == "":
        return jsonify({"error": "no file provided"}), 400
    if not _allowed_doc(file.filename):
        return jsonify({"error": "unsupported file type"}), 400

    safe_type = secure_filename(doc_type) or "Document"
    original = secure_filename(file.filename)
    stored_name = f"{safe_type}__{uuid.uuid4().hex[:10]}__{original}"
    folder = _student_doc_dir(student_id)
    file.save(os.path.join(folder, stored_name))

    return jsonify({"documents": _list_student_docs(student_id)})


@app.route("/api/students/<student_id>/documents/<filename>", methods=["DELETE"])
def api_student_document_delete(student_id, filename):
    if student_id not in ALL_STUDENTS_BY_ID:
        return jsonify({"error": "not found"}), 404
    user = current_user()
    if student_id in INCOMING_STUDENTS_BY_ID:
        allowed = auth_store.can_control(user, "student_exchange")
    else:
        allowed = any(
            auth_store.can_control(user, p) for p in ("universities", "student_exchange", "internships")
        )
    if not allowed:
        return jsonify({"error": "You don't have permission to modify student records. Ask an admin to grant you control of the relevant page."}), 403
    folder = _student_doc_dir(student_id)
    target = os.path.join(folder, secure_filename(filename))
    if os.path.isfile(target):
        os.remove(target)
    return jsonify({"documents": _list_student_docs(student_id)})


@app.route("/student-docs/<student_id>/<filename>")
def serve_student_doc(student_id, filename):
    return send_from_directory(_student_doc_dir(student_id), filename)


@app.route("/api/universities")
def api_universities():
    return jsonify(load_universities())


@app.route("/api/masters")
def api_masters():
    return jsonify({"universities": load_masters(), "version": _data_version})


@app.route("/api/masters/stats")
def api_masters_stats():
    unis = load_masters()
    programs = [p for u in unis for p in u["programs"]]
    direct = sum(1 for p in programs if "direct" in p.get("opportunity", "").lower())
    pathway = len(programs) - direct
    return jsonify({
        "universities": len(unis),
        "programs": len(programs),
        "countries": len({u["country"] for u in unis if u.get("country")}),
        "with_scholarship": sum(1 for p in programs if p.get("scholarship")),
        "with_ranking": sum(1 for p in programs if p.get("ranking")),
        "with_link": sum(1 for p in programs if p.get("courseLink")),
        "direct": direct,
        "pathway": pathway,
        "version": _data_version,
    })


@app.route("/api/stats")
def api_stats():
    unis = load_universities()
    return jsonify({
        "total":       len(unis),
        "active":      sum(1 for u in unis if u["status"] == "Active"),
        "renewal_due": sum(1 for u in unis if u["status"] == "Renewal Due"),
        "expired":     sum(1 for u in unis if u["status"] == "Expired"),
        "countries":   len({u["country"] for u in unis if u["country"]}),
        "version":     _data_version,
    })


@app.route("/api/stream")
def api_stream():
    """
    Persistent SSE connection — one per browser tab.
    Heartbeat every 20 s keeps the connection alive through proxies.
    """
    q = queue.Queue(maxsize=20)
    with _clients_lock:
        _clients.append(q)

    def generate():
        # Tell the client its SSE pipe is open
        yield f"event: connected\ndata: {_data_version}\n\n"
        try:
            while True:
                try:
                    yield q.get(timeout=20)
                except queue.Empty:
                    yield ": heartbeat\n\n"
        except GeneratorExit:
            pass
        finally:
            with _clients_lock:
                if q in _clients:
                    _clients.remove(q)

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control":     "no-cache",
            "X-Accel-Buffering": "no",      # disable nginx buffering
            "Connection":        "keep-alive",
        },
    )


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    start_watcher()
    # threaded=True  → each request (incl. SSE) runs in its own thread
    # use_reloader=False → prevents watchdog from being started twice
    # The Werkzeug debugger lets anyone who can reach the port run code, so it is
    # off unless you opt in:  OIA_HUB_DEBUG=1 python app.py
    # OIA_HUB_HOST=0.0.0.0 lets other computers on the network reach the Hub
    # (start_hub.bat sets this). Left unset, only this computer can open it.
    host = os.environ.get("OIA_HUB_HOST", "127.0.0.1")
    port = int(os.environ.get("OIA_HUB_PORT", "5000"))
    debug = os.environ.get("OIA_HUB_DEBUG") == "1"
    if debug and host not in ("127.0.0.1", "localhost"):
        print("[warning] Debug mode is never allowed on a shared network — turning it off.", flush=True)
        debug = False
    print(f"[hub] Serving on http://{host}:{port}", flush=True)
    app.run(host=host, port=port, debug=debug, threaded=True, use_reloader=False)