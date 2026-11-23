"""
Loads and normalizes student records from the OIA "Student List Abroad" Excel
export. The workbook has many sheets with slightly different column layouts
(different batches / exports over time), so this module maps each sheet's
header row to a common schema and de-duplicates by (name, UID).
"""

import os
import re
import hashlib
from datetime import datetime

import openpyxl


# Sheets that contain real per-student rosters (skips pivot/summary tabs like
# Sheet10/Sheet14, credit-only trackers with no university field like Sheet3,
# and name-only tabs like Sheet4 with no other usable fields).
#
# The 20-Jul-2026 "Student List Abroad" workbook consolidates every batch
# into one roll-up tab, "Final List (1-9)", which supersedes the individual
# Batch1..Batch10/Sheet1.../SE tabs (same students, same fields, already
# de-duplicated by OIA into one sheet). Loading only this sheet (instead of
# every batch tab) avoids re-introducing the duplicate/incomplete-record
# problem noted below for "Top 300", and matches what was asked for
# explicitly. "Top 300" (see load_students) is merged in separately since it
# has no Activity/Status/Course columns of its own.
SHEETS_TO_LOAD = [
    "Final List (1-9)",
]
# NOTE: earlier workbook exports used per-batch tabs ("Batch1".."Batch10",
# "Sheet1"/"Sheet2"/"Sheet6", "Credits Transfer", "SE", etc.) instead of a
# single consolidated sheet. If a future export drops "Final List (1-9)"
# and goes back to per-batch tabs, restore that list here.

# Activity values that count as an "internship-type" engagement. Shared with
# app.py (which imports this) so the Internships tab and the extra-workbook
# merge below always agree on what counts.
INTERNSHIP_ACTIVITIES = {"internship", "summer camp", "virtual internship", "summer school"}

# Additional standalone workbooks (per-sheet) to merge in alongside the main
# roster above. Each entry is (relative path from this file's data/ folder,
# [sheet names to load], default_activity, activity_filter).
#
# - Semester Exchange Consolidated Data: merged in full since its schema
#   lines up directly with HEADER_ALIASES and its "Overall Data" tab is a
#   roll-up of the Batch sheets (excluded for the same duplicate-avoidance
#   reason as SHEETS_TO_LOAD above).
#
# - Internship.xlsx (data/extra_sources): OIA's dedicated internship-roster
#   workbook. It shares the same batch layout as the main Student List
#   Abroad workbook but tracks some internship/summer-camp students (e.g.
#   several students at Ming Chi University of Technology, recorded here as
#   "Virtual Internship" — see ACTIVITY_CANONICAL above) that were never
#   entered into the main roster, so those hosts and students were silently
#   missing from the Internships tab. Only rows whose Activity is
#   internship-type are merged in (activity_filter) so this doesn't also
#   pull in the workbook's Masters/3+2/Semester-Exchange rows, which are
#   already tracked — and de-duplicated — via the main roster/Semester
#   Exchange workbook above.
EXTRA_WORKBOOKS = [
    (
        "extra_sources/Semester_Exchange_Consolidated_Data.xlsx",
        ["Overall Data "],
        "Semester Exchange",
        None,
    ),
    (
        "extra_sources/Internship.xlsx",
        [
            "Batch1", "Batch2", "Batch3", "Credits Transfer", "Sheet3",
            "Batch4", "Batch4.", "Batch5", "Batch4(1)", "Batch5.",
            "Batch 6", "Batch 7", "Batch 8", "Batch 9",
            "Sheet6", "Sheet4", "Sheet5", "Sheet2", "SE", "Sheet1",
        ],
        None,
        INTERNSHIP_ACTIVITIES,
    ),
]

# Maps normalized header text -> canonical field name.
HEADER_ALIASES = {
    "student name": "name",
    "name": "name",
    "uid no. /roll no.": "uid",
    "uid no /roll no": "uid",
    "student id": "uid",
    "roll number": "uid",
    "course": "course",
    "programm": "course",
    "course detail": "course",
    "program name": "program",
    "master program": "program",
    "activity": "activity",
    "acceptance letter": "acceptance_letter",
    "status": "status",
    "university & country": "university_country",
    "university name": "university",
    "university & country ": "university_country",
    "country": "country",
    "e-mail id": "email",
    "email id": "email",
    "e-mail id ": "email",
    "email id pu": "email",
    "phone number": "phone",
    "phone no.": "phone",
    "start date": "start_date",
    "end date": "end_date",
    "photo": "has_photo",
    "visa": "has_visa",
    "visa status": "has_visa",
    "certificate/gradecard": "has_certificate",
    "year": "year",
    "batch": "batch",
    "batch no": "batch",
    "scholarship": "scholarship",
    "cgpa": "cgpa",
    "application status": "status",
}


def _norm_header(cell):
    if cell is None:
        return ""
    return re.sub(r"\s+", " ", str(cell)).strip().lower()


def _clean(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    # Do not surface raw Excel error literals such as #REF! as real student
    # data. They represent a broken source formula, not a meaningful status.
    if text.upper() in {"#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}:
        return ""
    return text


def _split_university_country(raw):
    """'Nanyang Technological University, Singapore' -> (name, country)."""
    raw = _clean(raw)
    if not raw:
        return "", ""
    if "," in raw:
        parts = [p.strip() for p in raw.rsplit(",", 1)]
        if len(parts) == 2 and len(parts[1]) <= 30:
            return parts[0], parts[1]
    return raw, ""


def _make_id(name, uid, index):
    basis = f"{name}|{uid}|{index}"
    return "stu_" + hashlib.md5(basis.encode("utf-8")).hexdigest()[:10]


# Presidency University school inferred from the course/program code recorded
# per student. This is a best-effort mapping from course-code conventions
# seen in the roster (e.g. "B.Tech-CSE" -> Computer Science & Engineering);
# it is NOT an authoritative field from the source data, so it should be
# labeled as inferred wherever it's shown.
_SCHOOL_KEYWORDS = [
    # (school name, keywords to match against the normalized course/program text)
    ("School of Computer Science & Engineering",
     ["cse", "csd", "cst", "cit", "cai", "ise", "ist", "isr", "cce", "csg",
      "cbc", "ccs", "cei", "bca", "mca", "computer science"]),
    ("School of Engineering",
     ["ece", "eee", "mech", "civil", "pet", "b.tech", "btech", "m.tech", "mtech"]),
    ("School of Management",
     ["bba", "mba", "management"]),
    ("School of Commerce & Economics",
     ["bcom", "b.com", "commerce", "economics"]),
    ("School of Law",
     ["llb", "llm", "sol", "law"]),
    ("School of Design",
     ["design", "b.design"]),
    ("School of Media, Design & Arts",
     ["baj", "journalism", "media", "mass comm"]),
]


def infer_pu_school(course, program):
    """Best-effort PU School guess from course/program text. Returns '' if the
    course text already reads as a school name (used as-is by the caller) or
    nothing usable was found."""
    for raw in (course, program):
        text = (raw or "").strip()
        if not text:
            continue
        if text.lower().startswith("school of"):
            return re.sub(r"\band\b", "&", text, flags=re.IGNORECASE).strip()
        norm = re.sub(r"[^a-z0-9. ]", "", text.lower())
        for school, keywords in _SCHOOL_KEYWORDS:
            if any(kw in norm for kw in keywords):
                return school
    return ""


# Some source rows use activity labels that OIA wants folded into a
# different canonical label everywhere in the app (student profile, the
# Internships tab's host grouping/filtering, exports, etc.). Applied to
# every record as it's parsed, regardless of which workbook it came from.
ACTIVITY_CANONICAL = {
    "summer camp": "Virtual Internship",
}


def _canonicalize_activity(value):
    return ACTIVITY_CANONICAL.get((value or "").strip().lower(), value)


# Raw sheet-tab names that fall back into the "batch" field verbatim when a
# sheet has no explicit "Batch" column (see `get("batch") or sheet_name`
# below) — e.g. tabs literally named "Sheet6"/"Batch4."/"Batch4(1)" show up
# as-is on student profiles and the Internships tab, which isn't a
# consistent or meaningful label to OIA staff. Any tab name of the form
# "Batch<N>"/"Batch <N>"/"Batch<N>.'"/"Batch<N>(1)"/"Sheet<N>" is folded
# into the canonical "Batch <N>" so every batch shows one consistent label
# regardless of which tab a given student's row happened to load from.
# Anything that isn't a numbered batch/sheet tab (e.g. "Credits Transfer",
# "SE", "AY wise") is left as its own label rather than guessed at.
BATCH_LABEL_ALIASES = {}

_BATCH_NUMBER_RE = re.compile(r"^(?:batch|sheet)\s*-?\s*(\d+)")


def _canonicalize_batch(value):
    raw = (value or "").strip()
    if not raw:
        return raw
    key = raw.lower()
    if key in BATCH_LABEL_ALIASES:
        return BATCH_LABEL_ALIASES[key]
    m = _BATCH_NUMBER_RE.match(key)
    if m:
        return f"Batch {int(m.group(1))}"
    return raw


def _load_sheet_records(ws, sheet_name, default_activity=None, activity_filter=None):
    """Parses one worksheet into a list of raw student-record dicts using the
    HEADER_ALIASES mapping. Shared by the main roster and any extra
    workbooks merged in via EXTRA_WORKBOOKS.

    default_activity: used when the sheet has no "Activity" column at all
    (e.g. the Semester Exchange Consolidated workbook), so these rows can
    still merge with the same student's record from the main roster instead
    of loading as a separate "ghost" entry with a blank activity (which
    silently excludes them from activity-filtered views like /student-exchange).

    activity_filter: if given (a set of lowercased activity strings), rows
    whose Activity doesn't match are skipped entirely. Used so a workbook
    that mixes several engagement types (e.g. Internship.xlsx, which also
    has Masters/3+2/Semester-Exchange rows already tracked elsewhere) only
    contributes the rows it's actually meant to supplement."""
    rows = list(ws.iter_rows(values_only=True))

    header_idx = None
    col_map = {}
    for i, row in enumerate(rows):
        mapped = {}
        for c, cell in enumerate(row):
            key = _norm_header(cell)
            if key in HEADER_ALIASES:
                # Some source sheets contain duplicate logical headers (for
                # example, Batch 7/8/10 have both ``Status`` and a second
                # ``Status `` column used for a different tracker field).
                # Keep the FIRST occurrence so the primary roster field is
                # not silently overwritten by the secondary column.
                mapped.setdefault(HEADER_ALIASES[key], c)
        if "name" in mapped:
            header_idx = i
            col_map = mapped
            break
    if header_idx is None:
        return []

    records = []
    for row in rows[header_idx + 1:]:
        if row is None:
            continue

        def get(field):
            c = col_map.get(field)
            if c is None or c >= len(row):
                return ""
            return _clean(row[c])

        name = get("name")
        if not name:
            continue

        activity_value = get("activity") or (default_activity or "")
        activity_value = _canonicalize_activity(activity_value)
        # A few old workbook tabs contain a second, differently shaped table
        # without a repeated header. When read using the first table's schema,
        # its Acceptance Letter ("Yes") lands in the Activity field and the
        # university lands in Status. "Yes"/"No" are not valid OIA activity
        # labels, so discard those malformed rows instead of creating visibly
        # corrupted student records.
        if activity_value.strip().lower() in {"yes", "no"} and default_activity is None:
            continue
        if activity_filter is not None and activity_value.strip().lower() not in activity_filter:
            continue

        uid = get("uid")
        uni_country_raw = get("university_country")
        uni_name, uni_country = _split_university_country(uni_country_raw)
        if not uni_name:
            uni_name = get("university")

        # Historical exports contain two kinds of contact-column corruption:
        # (1) some sheets omitted the Country header, shifting Country -> Email
        #     and Email -> Phone; and
        # (2) some consolidated copies have Country in the email column while
        #     the real email/phone live elsewhere.
        # Repair both patterns before records are merged.
        email = get("email")
        phone = get("phone")
        country = get("country") or uni_country

        def _is_email(value):
            return bool(re.search(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", (value or "").strip()))

        def _is_phone(value):
            raw = re.sub(r"[^0-9]", "", (value or "").strip())
            return 7 <= len(raw) <= 15 and not _is_email(value)

        # Missing-country-header layout: email currently contains Country and
        # phone currently contains Email. The actual phone is the next cell.
        if "country" not in col_map:
            if not _is_email(email) and _is_email(phone):
                country = email or country
                email = phone
                phone_col = col_map.get("phone")
                next_phone = _clean(row[phone_col + 1]) if phone_col is not None and phone_col + 1 < len(row) else ""
                if next_phone:
                    phone = next_phone

        # Consolidated-copy layout: Country was copied into the email field.
        # If the value is clearly not an email, preserve it as Country rather
        # than allowing a later duplicate record to corrupt the real email.
        if email and not _is_email(email):
            looks_like_country = (
                not _is_phone(email)
                and len(email) <= 40
                and (
                    email.lower() in {
                        "usa", "uk", "united kingdom", "usa / new zealand",
                        "canada", "cananda", "australia", "germany", "france",
                        "italy", "spain", "taiwan", "malaysia", "thailand",
                        "ireland", "new zealand", "latvia", "turkey", "singapore",
                        "south korea", "japan", "china", "netherlands", "switzerland"
                    }
                    or email.lower().startswith(("taiwan (", "uk (", "usa ("))
                )
            )
            if looks_like_country:
                country = country or email
                email = ""

        # If phone is accidentally an email, keep it out of the phone field.
        # A subsequent duplicate sheet may contain the real number, and the
        # merge logic below will prefer a valid phone over this bad value.
        if _is_email(phone):
            phone = ""

        record = {
            "name": name,
            "uid": uid,
            "course": get("course"),
            "program": get("program"),
            "pu_school": infer_pu_school(get("course"), get("program")),
            "activity": activity_value,
            "acceptance_letter": get("acceptance_letter"),
            "status": get("status") or "On record",
            "university": uni_name,
            "country": country,
            "email": email,
            "phone": phone,
            "start_date": get("start_date"),
            "end_date": get("end_date"),
            "has_photo": get("has_photo"),
            "has_visa": get("has_visa"),
            "has_certificate": get("has_certificate"),
            "year": get("year"),
            # Only infer a batch from an actual Batch-named worksheet.
            # Generic tabs such as Sheet1/Sheet6 are consolidated copies and
            # must not be mistaken for Batch 1/Batch 6.  An explicit Batch
            # column always wins.
            "batch": _canonicalize_batch(
                get("batch") or (sheet_name if re.match(r"^batch\s*-?\s*\d+", sheet_name, re.IGNORECASE) else "")
            ),
            "scholarship": get("scholarship"),
            "cgpa": get("cgpa"),
        }
        records.append(record)
    return records


# Sheets within the MAIN workbook that have no "Activity" column of their own
# (e.g. "SE" is a credits-transfer tracker for Semester Exchange students) —
# same reasoning/fix as EXTRA_WORKBOOKS' default_activity above.
SHEET_DEFAULT_ACTIVITY = {
    "SE": "Semester Exchange",
}


def load_students(xlsx_path):
    if not os.path.exists(xlsx_path):
        return []

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    records = []

    for sheet_name in SHEETS_TO_LOAD:
        if sheet_name not in wb.sheetnames:
            continue
        records.extend(_load_sheet_records(
            wb[sheet_name], sheet_name, SHEET_DEFAULT_ACTIVITY.get(sheet_name)
        ))

    # "Top 300" (QS/THE top-300-ranked-university placements) has no
    # Activity/Status/Course columns of its own, and its roll numbers are
    # ~98% already present in "Final List (1-9)" (it's a ranking-sorted
    # view of the same roster, not a distinct one). Merging it in generically
    # would create a near-duplicate "ghost" row for every student already
    # loaded above (same person, blank activity, so it wouldn't merge with
    # their real record). Instead, only students whose Roll Number isn't
    # already in the loaded roster get added from this sheet, so the sheet
    # still contributes its one genuinely new student without duplicating
    # the other ~59.
    if "Top 300" in wb.sheetnames:
        known_uids = {r["uid"].strip().lower() for r in records if r.get("uid")}
        for rec in _load_sheet_records(wb["Top 300"], "Top 300", default_activity=""):
            if rec["uid"].strip().lower() not in known_uids:
                records.append(rec)
                known_uids.add(rec["uid"].strip().lower())

    base_dir = os.path.dirname(xlsx_path)
    for rel_path, sheet_names, default_activity, activity_filter in EXTRA_WORKBOOKS:
        extra_path = os.path.join(base_dir, rel_path)
        if not os.path.exists(extra_path):
            continue
        extra_wb = openpyxl.load_workbook(extra_path, data_only=True)
        for sheet_name in sheet_names:
            if sheet_name not in extra_wb.sheetnames:
                continue
            records.extend(_load_sheet_records(
                extra_wb[sheet_name], sheet_name, default_activity, activity_filter
            ))

    # De-duplicate by (normalized name, uid, activity); merge records while
    # preferring values that pass basic field validation. This is important
    # because older consolidated tabs can contain shifted/corrupted copies of
    # the same student (for example Country in Email).
    # Activity is part of the key because the same student can genuinely
    # have multiple distinct engagements on record under one UID (e.g. a
    # completed Semester Exchange *and* a separate later Masters
    # application) — those must stay as separate rows, not collapse into
    # one and silently drop whichever activity loaded first.
    merged = {}
    order = []
    for rec in records:
        key = (
            rec["name"].strip().lower(),
            rec["uid"].strip().lower(),
            (rec.get("activity") or "").strip().lower(),
        )
        if key not in merged:
            merged[key] = rec
            order.append(key)
        else:
            existing = merged[key]

            def _email_ok(v):
                return bool(re.search(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", (v or "").strip()))

            def _phone_ok(v):
                if not v or _email_ok(v):
                    return False
                digits = re.sub(r"[^0-9]", "", str(v))
                return 7 <= len(digits) <= 15

            for k, v in rec.items():
                if not v:
                    continue
                old = existing.get(k, "")
                if not old:
                    existing[k] = v
                    continue

                # Never replace a valid email/phone with a lower-quality copy.
                if k == "email" and _email_ok(v) and not _email_ok(old):
                    existing[k] = v
                elif k == "phone" and _phone_ok(v) and not _phone_ok(old):
                    existing[k] = v
                # The parser uses this placeholder only when no source status
                # exists. If a duplicate sheet has a real status, keep it.
                elif k == "status" and old.strip().lower() == "on record" and v.strip().lower() != "on record":
                    existing[k] = v
                # A malformed email may also have leaked into Country in an
                # older copy; prefer a normal country string.
                elif k == "country" and _email_ok(old) and not _email_ok(v):
                    existing[k] = v

    final = []
    for i, key in enumerate(order):
        rec = merged[key]
        rec["id"] = _make_id(rec["name"], rec["uid"], i)
        final.append(rec)

    return final


# Known abbreviations / typos in the raw Excel university text that won't
# tokenize-match the partner university's full name automatically.
# Keys and values are both matched after normalization, so partial names work.
UNIVERSITY_ALIASES = {
    "uwe bristol": "university of west of england bristol",
    "vcu": "virginia commonwealth university",
    "smu lyle school of engineering": "southern methodist university",
    "queen mary": "queen mary university of london",
    "la trobe": "la trobe university",
    "nottingam trent": "nottingham trent university",
    "sandiego": "san diego",
    "unsw": "university of new south wales",
    "usm": "universiti sains malaysia",
    "iowa state university": "iowa state university of science and technology",
    "mcut": "ming chi university of technology",
    "iqs": "iqs sm universitat ramon llull",
}


def _normalize(text):
    """Lowercase and replace any non-alphanumeric run with a single space, so
    punctuation (dashes, en-dashes, apostrophes, parentheses) never fuses two
    words together (e.g. 'Wisconsin-La' must not become 'wisconsinla')."""
    text = (text or "").lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return UNIVERSITY_ALIASES.get(text, text)


def _similar(a, b, threshold=0.86):
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio() >= threshold


# Known duplicate spellings/typos/abbreviations found in the raw "University"
# text for Semester Exchange records specifically — these are real host
# universities that mostly aren't in the Masters partner list, so they need
# their own canonical-name grouping rather than being matched against it.
EXCHANGE_UNIVERSITY_CANONICAL = {
    "aalen univeristy": "Aalen University",
    "aalen university": "Aalen University",
    "fenerbahce university": "Fenerbahçe University",
    "fenerbahçe university": "Fenerbahçe University",
    "hochschule mannheim": "Hochschule Mannheim",
    "hochschule mannheim hochschule mannheim": "Hochschule Mannheim",
    "hochschule mannheim university of applied sciences": "Hochschule Mannheim",
    "iqs": "IQS SM_Universitat Ramon Llull",
    "isep": "Institut supérieur d’électronique de Paris (ISEP)",
    "mcut": "Ming Chi University of Technology",
    "national sun yat sen university": "National Sun Yat-Sen University",
    "st petersburg university": "St. Petersburg State University (SPbU)",
    "st petersburg state university spbu": "St. Petersburg State University (SPbU)",
    "tamakang university": "Tamkang University",
    "tamkang university": "Tamkang University",
    "usm": "Universiti Sains Malaysia",
    "uthm": "Universiti Tun Hussein Onn Malaysia",
    "universiti tun hussein onn malaysia": "Universiti Tun Hussein Onn Malaysia",
    "universite d orleans": "University of Orléans",
    "university of orleans": "University of Orléans",
    "uwe bristol": "University of West of England Bristol",
}


def canonicalize_exchange_university(name):
    """Groups near-duplicate spellings of the same host university under one
    canonical display name, for Semester Exchange records specifically."""
    if not name:
        return None
    import unicodedata
    ascii_name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    key = re.sub(r"[^a-z0-9]+", " ", ascii_name.lower())
    key = re.sub(r"\s+", " ", key).strip()
    return EXCHANGE_UNIVERSITY_CANONICAL.get(key, name.strip())


def find_students_for_university(students, university_name):
    """Loose match between a partner university's name and the free-text
    university field recorded per student: substring match, token-subset
    match (ignoring generic words), and a fuzzy-typo fallback."""
    target = _normalize(university_name)
    if not target:
        return []
    stop = {"university", "of", "the", "college", "institute", "school", "and"}
    target_tokens = set(target.split()) - stop

    matches = []
    for s in students:
        uni = _normalize(s.get("university"))
        if not uni:
            continue

        if target in uni or uni in target:
            matches.append(s)
            continue

        uni_tokens = set(uni.split()) - stop
        if target_tokens and uni_tokens and (
            target_tokens.issubset(uni_tokens) or uni_tokens.issubset(target_tokens)
        ):
            matches.append(s)
            continue

        if _similar(target, uni):
            matches.append(s)

    return matches
