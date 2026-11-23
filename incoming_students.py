"""
Loader for the Inbound Exchange dataset — students coming FROM partner
universities TO Parul University. Mirrors the shape of students_data.py's
outbound STUDENTS list closely enough to reuse the same detail-page /
document-gallery UI, but is kept as a fully separate list so it never mixes
with outbound analytics.
"""
import hashlib
import os
import re
import unicodedata
from datetime import datetime

import pandas as pd


# The source sheet writes the same host institution a dozen different ways
# ("HTWG, Konstanz, Germany" / "Konstanz University of Applied Sciences or
# HTWG, Germany" / ...). This maps a normalized key -> (canonical name, country).
# NOTE on "Morocco, France" (1 student, row 17 in the source sheet): the raw
# label doesn't cleanly say which institution this is. It's mapped below as
# an ISEP Morocco-campus record (ISEP runs a Rabat/Morocco campus and is
# otherwise the dominant partner in this sheet) — treat that single record
# as a best-effort guess, not a confirmed mapping.
_CANON_MAP = {
    "academy of national economy and public administration ranepa russia": (
        "Academy of National Economy and Public Administration (RANEPA)", "Russia"),
    "cy tech france": ("CY Tech", "France"),
    "cy techfrance": ("CY Tech", "France"),
    "ece ecole centrale d electronique paris": ("ECE Paris (École d'Ingénieurs)", "France"),
    "ece engg france": ("ECE Paris (École d'Ingénieurs)", "France"),
    "ece enggfrance": ("ECE Paris (École d'Ingénieurs)", "France"),
    "epita": ("EPITA", "France"),
    "htwg konstanz germany": ("HTWG Konstanz – University of Applied Sciences", "Germany"),
    "konstanz university of applied sciences or htwg germany": (
        "HTWG Konstanz – University of Applied Sciences", "Germany"),
    "isep france": ("ISEP", "France"),
    "morocco france": ("ISEP (Morocco Campus)", "Morocco"),
    "polytech orleans france": ("Polytech Orléans", "France"),
    "polytech tours university of tours": ("Polytech Tours (University of Tours)", "France"),
    "polytechnic university of tours france": ("Polytech Tours (University of Tours)", "France"),
    "saint petersburg state university russia": ("Saint Petersburg State University", "Russia"),
    "saint petersburg state universityrussia": ("Saint Petersburg State University", "Russia"),
    "th aschaffenburg": ("Technische Hochschule Aschaffenburg", "Germany"),
    "thws": ("Technische Hochschule Würzburg-Schweinfurt (THWS)", "Germany"),
    "tallinn university of technology estonia": ("Tallinn University of Technology", "Estonia"),
    "technikum wien": ("FH Technikum Wien", "Austria"),
    "university of applied management studies": ("University of Applied Management Studies", "Germany"),
}


def _key(raw: str) -> str:
    ascii_name = unicodedata.normalize("NFKD", raw or "").encode("ascii", "ignore").decode("ascii")
    key = re.sub(r"[^a-z0-9]+", " ", ascii_name.lower())
    return re.sub(r"\s+", " ", key).strip()


def canonicalize_incoming_university(raw: str):
    """Returns (canonical_name, country) for a raw host-university string
    from the incoming-student sheet, falling back to a best-effort comma
    split (institution, ..., country) when the string isn't in the map."""
    raw = (raw or "").strip()
    if not raw:
        return None, ""
    mapped = _CANON_MAP.get(_key(raw))
    if mapped:
        return mapped
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) >= 2:
        return ", ".join(parts[:-1]), parts[-1]
    return raw, ""


def _clean(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "*", "-") else s


def _fmt_date(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d %b %Y")
    s = str(val).strip()
    if not s or s.lower() == "nat":
        return ""
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%d %b %Y")
    except Exception:
        return s


def _status_for(start_val, end_val) -> str:
    try:
        start = pd.to_datetime(start_val, dayfirst=True)
        end = pd.to_datetime(end_val, dayfirst=True)
    except Exception:
        return "On record"
    if pd.isna(start) or pd.isna(end):
        return "On record"
    today = pd.Timestamp(datetime.now().date())
    if today < start:
        return "Upcoming"
    if today > end:
        return "Completed"
    return "Ongoing"


INCOMING_EXCEL = os.path.join(
    os.path.dirname(__file__), "data", "extra_sources", "Incoming_Student_Data.xlsx"
)


def load_incoming_students(path: str = INCOMING_EXCEL) -> list[dict]:
    """Parses the 'Student Data' sheet (two-row header: main header on row 2,
    'From'/'To' sub-header for Period on row 3) into a flat list of student
    dicts shaped like students_data.py's outbound records, so the existing
    student-detail page / document gallery work unmodified."""
    if not os.path.isfile(path):
        return []

    try:
        raw = pd.read_excel(path, sheet_name="Student Data", header=None)
    except Exception as exc:
        print(f"[incoming-students] Failed to read {path}: {exc}", flush=True)
        return []

    students = []
    for _, r in raw.iloc[3:].iterrows():
        name = _clean(r.get(1))
        if not name:
            continue

        gender      = _clean(r.get(2))
        passport    = _clean(r.get(3))
        email       = _clean(r.get(4))
        ay          = _clean(r.get(5))
        period_from = r.get(6)
        period_to   = r.get(7)
        uni_raw     = _clean(r.get(8))
        department  = _clean(r.get(9))

        canonical_uni, country = canonicalize_incoming_university(uni_raw)
        start_date = _fmt_date(period_from)
        end_date   = _fmt_date(period_to)

        sid = "in_" + hashlib.md5(f"{name}|{uni_raw}|{email}".encode("utf-8")).hexdigest()[:10]

        students.append({
            "id":                 sid,
            "name":               name,
            "gender":             gender,
            "uid":                passport,
            "email":              email,
            "phone":              "",
            "batch":              ay,
            "year":               ay,
            "university":         canonical_uni,
            "university_raw":     uni_raw,
            "country":            country,
            "course":             department,
            "program":            department,
            "activity":           "Inbound Exchange",
            "acceptance_letter":  "",
            "scholarship":        "",
            "start_date":         start_date,
            "end_date":           end_date,
            "status":             _status_for(period_from, period_to),
        })
    return students
