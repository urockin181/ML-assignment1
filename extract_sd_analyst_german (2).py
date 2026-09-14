#!/usr/bin/env python3
"""
Daily extract: clean the AUDI/VW hiring workbook and keep only rows for the
Service Desk Analyst - German role family, tier A4 only:

    A4  "Associate Service Desk Analyst - German"  (entry tier -- KEPT)
    A5  "Service Desk Analyst - German"             (senior tier -- EXCLUDED)

The A5 "Senior" tier is recognized (so it gets a clean "excluded" log line)
but dropped everywhere it appears, per TARGET_TIERS at the top of this file.
Flip that back to {"A4", "A5"} to include the senior tier again.

Each sheet in the source workbook names this role family differently, so the
script keeps one alias table (ROLE_ALIASES_BY_SHEET below) mapping every
known spelling, on every sheet, to its tier -- and rewrites matched cells to
a clean, canonical spelling for that sheet along the way.

What it does
------------
- Deletes every sheet NOT listed in SHEETS_TO_KEEP outright (edit that list
  to change what survives). The sheets that do stay keep their original
  order and column layout otherwise.
- Drops every row, in every role-bearing sheet, that isn't part of this role
  family at all, AND every row that's the family's A5 tier (out of scope --
  see TARGET_TIERS).
- Trims whitespace / non-breaking spaces everywhere, fixes header typos,
  normalizes status values and dates, exactly as the earlier full-workbook
  cleaner did.
- Fixes a specific, verified inconsistency: in Offering Sheet, "Skills" is
  sometimes copy-pasted from the wrong tier (e.g. Grade A5 with the A4
  "Associate..." label). Grade is the reliable field there, so it's also
  what decides A4-vs-A5 for that sheet's tier filter, not the Skills text --
  otherwise a mislabeled A5 row would slip through as "A4" by its text alone.
- Anything it can't confidently resolve is left alone and reported in the
  log -- including a loud warning for any role-like text that doesn't match
  a known alias, so a new typo or naming drift doesn't silently vanish.
- Writes a timestamped .log.txt next to the output with everything it did,
  since the output workbook itself must stay identical in shape to the input.

Usage
-----
    python extract_sd_analyst_german.py
        Cleans INPUT_PATH and writes to OUTPUT_PATH, both edited into this
        file below (look for '>>> EDIT THIS <<<'). Nothing to type each day.

    python extract_sd_analyst_german.py INPUT.xlsx OUTPUT.xlsx
        Overrides both paths for this one run, without touching the file --
        handy for testing against a different export.

The input file is never modified -- only OUTPUT_PATH is written.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# >>> EDIT THIS <<<
# Full path to the raw daily export. Point this at wherever that file lands
# each day (e.g. a synced OneDrive/Downloads folder) and just re-run the
# script -- no command-line arguments needed.
#
# Windows example:  r"C:\Users\yourname\Documents\AUDI_VW_Final_Funnel.xlsx"
# Mac/Linux example: "/Users/yourname/Documents/AUDI_VW_Final_Funnel.xlsx"
# --------------------------------------------------------------------------
INPUT_PATH = r"PASTE_THE_FULL_PATH_TO_TODAYS_EXPORT_HERE.xlsx"

# Full path for the cleaned file. This is NOT optional to think about --
# pick exactly where you want it (e.g. the folder Power BI reads from).
# Windows example: r"C:\Users\yourname\Documents\SD_Analyst_German_clean.xlsx"
OUTPUT_PATH = r"PASTE_THE_FULL_PATH_FOR_THE_CLEANED_OUTPUT_HERE.xlsx"

# Only these sheets survive into the output; every other sheet in the source
# workbook is deleted outright (not just left unfiltered). Edit this list
# directly if a different set is needed -- it's the one place that decides
# what ends up in the file.
SHEETS_TO_KEEP = [
    "Candidate Tracker",
    "Offering Sheet",
    "Positions",
    "Roles Overview",
    "SD Analyst Overview",
]

# --------------------------------------------------------------------------
# Role family alias table.
#
# Key: normalized text (see normalize_role_text) as it's been seen spelled,
# on any sheet, ever. Value: (tier, canonical spelling for THIS sheet).
# Add a new row here the day a new spelling shows up -- the log will tell
# you when one doesn't match anything below.
# --------------------------------------------------------------------------

CANDIDATE_SHEET = "Candidate Tracker"

# Which tiers of the role family to actually keep. Both tiers are still
# recognized below (so a real A5 row gets a clean, informative "excluded"
# log line instead of the scarier "unrecognized role" warning) -- this is
# the one knob that controls what ends up in the output. Put "A5" back in
# this set to include the senior tier again.
TARGET_TIERS = {"A4"}

ROLE_ALIASES_BY_SHEET = {
    "Candidate Tracker": {
        "column": "Role Applied",
        "aliases": {
            "sd analyst german": ("A4", "SD Analyst German"),
            "snr sd analyst german": ("A5", "Snr SD Analyst German"),
        },
    },
    "Positions": {
        "column": "Role",
        "aliases": {
            "german service desk agent": ("A4", "German Service Desk Agent"),
            "german senior service desk agent": ("A5", "German Senior Service Desk Agent"),
        },
    },
    "Offering Sheet": {
        "column": "Skills",
        # Grade is the field recruiters fill in deliberately here (Skills is
        # often stale copy-paste -- see fix_offering_sheet_skills), so for
        # this sheet only, an explicit Grade column wins the tier decision.
        "grade_overrides_tier": True,
        "aliases": {
            "associate service desk analyst - german": ("A4", "Associate Service Desk Analyst - German"),
            "service desk analyst - german": ("A5", "Service Desk Analyst - German"),
        },
    },
    "Roles Overview": {
        "column": "Role",
        "aliases": {
            "sd analayst": ("A4", "SD Analyst"),
            "sd analyst": ("A4", "SD Analyst"),
            "snr sd analyst": ("A5", "Snr SD Analyst"),
        },
    },
    # Already scoped entirely to this role family in practice, but still
    # worth running through the same tier filter/typo-fix for consistency.
    "SD Analyst Overview": {
        "column": "Role",
        "aliases": {
            "sd analayst": ("A4", "SD Analyst"),
            "sd analyst": ("A4", "SD Analyst"),
            "snr sd analyst": ("A5", "Snr SD Analyst"),
        },
    },
}

# A soft signal that a dropped value MIGHT be a misspelled member of this
# family rather than a genuinely different role -- triggers a loud log line
# instead of a silent drop.
SOFT_SIGNAL_WORDS = ("service desk", "sd analy", "sd agent")

# Offering Sheet: which Skills text a kept row should have, purely as a
# function of Grade (Grade is the field recruiters fill in deliberately;
# Skills is often stale copy-paste -- see the log for confirmed cases).
GRADE_TO_SKILLS = {
    "A4": "Associate Service Desk Analyst - German",
    "A5": "Service Desk Analyst - German",
}

STATUS_COLUMNS = {"HR Screening", "Technical", "Ops", "Accepted", "Offer Status"}
STATUS_SYNONYMS = {
    "done": "Done", "pass": "Passed", "passed": "Passed", "fail": "Failed", "failed": "Failed",
    "ncns": "NCNS", "no show": "NCNS", "reschedule": "Rescheduled", "rescheduled": "Rescheduled",
    "cancel": "Cancelled", "canceled": "Cancelled", "cancelled": "Cancelled",
    "accept": "Accepted", "accepted": "Accepted", "reject": "Rejected", "rejected": "Rejected",
    "pending": "Pending", "in progress": "In Progress", "not started": "Not Started",
    "joined": "Joined", "internal": "Internal",
}
PROJECT_SYNONYMS = {"audi": "AUDI", "vw": "VW", "volkswagen": "VW"}
NON_HIRE_STATUSES = {"Rejected", "Cancelled", "NCNS", "Failed"}
NAME_LIKE_COLUMNS = {"Name", "Recruiter", "Candidate Full Name"}
DATE_FORMATS = ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y")
MONTHS_PATTERN = re.compile(r"^(\d+)\s*months?$", re.IGNORECASE)

# Header spelling fixes, applied by sheet (typos seen in the source file).
HEADER_FIXES = {
    "Positions": {"Gab": "Gap"},
    "Roles Overview": {"1st of xNov": "1st of Nov", "1st of sep": "1st of Sep"},
}

STRAY_PUNCTUATION = set("`~!@#$%^&*_+-=|\\")


# --------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------

def normalize_whitespace(value: str) -> str:
    value = unicodedata.normalize("NFKC", value)
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_role_text(value) -> str:
    if not isinstance(value, str):
        return ""
    return normalize_whitespace(value).lower()


def smart_title_case(value: str) -> str:
    words = value.split(" ")
    fixed = []
    for w in words:
        upper = w.upper()
        if upper in {"SD", "HR", "UX", "XMO", "IT", "QA", "VW", "BU", "CIS"}:
            fixed.append(upper)
        elif "-" in w:
            fixed.append("-".join(p.capitalize() for p in w.split("-")))
        else:
            fixed.append(w.capitalize())
    return " ".join(fixed)


def normalize_status(value: str) -> str:
    key = value.strip().lower()
    return STATUS_SYNONYMS.get(key, smart_title_case(value))


def normalize_project(value: str) -> str:
    key = value.strip().lower()
    return PROJECT_SYNONYMS.get(key, smart_title_case(value))


def add_months(dt: datetime, n: int) -> datetime:
    month_index = dt.month - 1 + n
    year = dt.year + month_index // 12
    month = month_index % 12 + 1
    days_in_month = [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
                      31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    day = min(dt.day, days_in_month[month - 1])
    return dt.replace(year=year, month=month, day=day)


def parse_month_name(text: str):
    for fmt in ("%B", "%b"):
        try:
            return datetime.strptime(text.strip().title(), fmt).month
        except ValueError:
            continue
    return None


def parse_date(value):
    if isinstance(value, datetime):
        return value
    if not isinstance(value, str):
        return None
    text = normalize_whitespace(value)
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


# --------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------

class Log:
    def __init__(self):
        self.lines: list[str] = []
        self.fixed_count = 0
        self.dropped_count = 0
        self.warnings: list[str] = []

    def note(self, text: str):
        self.lines.append(text)

    def fixed(self, text: str):
        self.fixed_count += 1
        self.lines.append("[fixed]   " + text)

    def dropped(self, sheet: str, n: int):
        self.dropped_count += n
        self.lines.append(f"[dropped] {sheet}: removed {n} row(s) outside the role family")

    def excluded(self, sheet: str, n: int, tier: str):
        self.dropped_count += n
        self.lines.append(f"[excluded] {sheet}: removed {n} row(s) -- tier {tier} is out of scope "
                           f"(TARGET_TIERS = {sorted(TARGET_TIERS)})")

    def warn(self, text: str):
        self.warnings.append(text)
        self.lines.append("[REVIEW]  " + text)

    def render(self) -> str:
        header = [
            f"SD Analyst German / Associate SD Analyst German -- daily extract log",
            f"Run at: {datetime.now().isoformat(timespec='seconds')}",
            f"Cells normalized: {self.fixed_count}   Rows dropped: {self.dropped_count}   "
            f"Items needing review: {len(self.warnings)}",
            "-" * 78,
        ]
        return "\n".join(header + self.lines) + "\n"


# --------------------------------------------------------------------------
# Cell-level cleaning (applied to every sheet, every cell, before filtering)
# --------------------------------------------------------------------------

def sheet_key(name: str) -> str:
    """Sheet tab names in this workbook have shown up with stray trailing
    whitespace (e.g. 'Roles Overview '); match config by the trimmed name
    so a sloppy tab name never silently breaks filtering."""
    return name.strip() if isinstance(name, str) else name


def clean_headers(ws, log: Log):
    fixes = HEADER_FIXES.get(sheet_key(ws.title), {})
    headers = []
    for cell in ws[1]:
        raw = cell.value
        if isinstance(raw, str):
            clean = normalize_whitespace(raw)
            if clean in fixes:
                log.fixed(f"{ws.title}!{cell.coordinate}: header '{clean}' -> '{fixes[clean]}'")
                clean = fixes[clean]
            if clean != raw:
                cell.value = clean if clean else None
                log.fixed_count += 1
            headers.append(clean)
        else:
            headers.append(raw)
    return headers


def clean_cell_generic(cell, header, log: Log):
    """Whitespace/casing cleanup for one data cell. Returns the (possibly
    updated) value. Role and date columns are handled separately."""
    value = cell.value
    if cell.data_type == "f" or not isinstance(value, str):
        return value

    cleaned = normalize_whitespace(value)
    if not cleaned:
        if value != cleaned:
            log.fixed_count += 1
        return None

    if header in NAME_LIKE_COLUMNS:
        cleaned = smart_title_case(cleaned)
    elif header == "Project":
        cleaned = normalize_project(cleaned)
    elif header in STATUS_COLUMNS:
        cleaned = normalize_status(cleaned)
    elif header == "Grade":
        cleaned = cleaned.upper()

    if cleaned != value:
        log.fixed_count += 1
    return cleaned


def resize_tables_to_fit(ws, log: Log):
    """An Excel Table (ListObject) needs two things kept in sync with the
    sheet, and openpyxl does neither automatically:

    1. Its declared ref -- ws.delete_rows() doesn't shrink it, so after
       filtering it keeps claiming the old (now too-large) row range.
    2. Its cached per-column names -- a completely separate copy of the
       header text baked into the table definition, unaffected by editing
       the header cells themselves.

    Either mismatch alone is exactly what Excel's "we found a problem with
    some content" recovery prompt is warning about; both are fixed here so
    the table entry in the file matches what's actually on the sheet."""
    for table in list(ws.tables.values()):
        min_col, min_row, max_col, _old_max_row = range_boundaries(table.ref)
        new_max_row = max(ws.max_row, min_row)  # keep at least the header row
        new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"
        if new_ref != table.ref:
            log.fixed(f"{ws.title}: table '{table.name}' resized {table.ref} -> {new_ref} "
                      "(prevents the 'we found a problem with some content' Excel warning)")
            table.ref = new_ref
            if table.autoFilter is not None:
                table.autoFilter.ref = new_ref

        for i, col in enumerate(table.tableColumns):
            header_cell = ws.cell(row=min_row, column=min_col + i)
            actual = header_cell.value
            if actual and col.name != actual:
                log.fixed(f"{ws.title}: table '{table.name}' column name '{col.name}' -> '{actual}' "
                          "(was out of sync with the actual header cell)")
                col.name = actual


def clear_stray_punctuation_rows(ws, headers, log: Log):
    """A row whose only non-blank cell is a single punctuation character
    (a stray keystroke, e.g. a lone backtick) gets that cell cleared."""
    for row in ws.iter_rows(min_row=2):
        values = [c.value for c in row]
        non_blank = [v for v in values if v not in (None, "")]
        if len(non_blank) == 1 and isinstance(non_blank[0], str) and non_blank[0].strip() in STRAY_PUNCTUATION:
            for cell in row:
                if cell.value == non_blank[0]:
                    log.fixed(f"{ws.title}!{cell.coordinate}: cleared stray character {non_blank[0]!r}")
                    cell.value = None


# --------------------------------------------------------------------------
# Role-family filtering
# --------------------------------------------------------------------------

def filter_role_rows(ws, headers, log: Log):
    """Drop rows outside the role family for sheets that have a role column.
    Rewrites the matched cell to the sheet's canonical spelling. No-op for
    sheets not listed in ROLE_ALIASES_BY_SHEET."""
    config = ROLE_ALIASES_BY_SHEET.get(sheet_key(ws.title))
    if config is None:
        return None  # not a role-bearing sheet -- keep everything

    col_index = {h: i + 1 for i, h in enumerate(headers) if h}
    role_col = col_index.get(config["column"])
    if role_col is None:
        log.warn(f"{ws.title}: expected column '{config['column']}' not found -- sheet left unfiltered")
        return None
    grade_col = col_index.get("Grade") if config.get("grade_overrides_tier") else None

    aliases = config["aliases"]
    rows_to_drop = []
    out_of_scope: list[int] = []
    excluded_by_tier: dict[str, int] = {}
    tiers = {}  # row_num -> tier, for downstream use (Offering Sheet Skills fix)

    for row in ws.iter_rows(min_row=2):
        row_num = row[0].row
        if all(c.value in (None, "") for c in row):
            continue  # fully blank rows are handled elsewhere
        cell = ws.cell(row=row_num, column=role_col)
        text = cell.value
        key = normalize_role_text(text)
        if not key:
            rows_to_drop.append(row_num)
            continue
        match = aliases.get(key)
        if match is None:
            if any(sig in key for sig in SOFT_SIGNAL_WORDS):
                log.warn(
                    f"{ws.title}!{cell.coordinate}: '{text}' looks role-family-ish but matches no "
                    "known alias -- ROW DROPPED, please check for a new typo/spelling"
                )
            rows_to_drop.append(row_num)
            continue
        tier, canonical = match

        if grade_col is not None:
            grade_val = ws.cell(row=row_num, column=grade_col).value
            if grade_val in ("A4", "A5") and grade_val != tier:
                tier = grade_val  # Grade is authoritative for this sheet -- see config comment

        if tier not in TARGET_TIERS:
            out_of_scope.append(row_num)
            excluded_by_tier[tier] = excluded_by_tier.get(tier, 0) + 1
            continue

        tiers[row_num] = tier
        if isinstance(text, str) and normalize_whitespace(text) != canonical:
            log.fixed(f"{ws.title}!{cell.coordinate}: '{text.strip()}' -> '{canonical}'")
        cell.value = canonical

    if rows_to_drop:
        log.dropped(ws.title, len(rows_to_drop))
    for tier, n in sorted(excluded_by_tier.items()):
        log.excluded(ws.title, n, tier)

    for row_num in reversed(sorted(rows_to_drop + out_of_scope)):
        ws.delete_rows(row_num)

    return tiers


def fix_offering_sheet_skills(ws, headers, log: Log):
    """Grade is the reliable field in Offering Sheet -- re-derive Skills
    from it for every remaining row, fixing stale copy-pasted tier labels."""
    col_index = {h: i + 1 for i, h in enumerate(headers) if h}
    skills_col = col_index.get("Skills")
    grade_col = col_index.get("Grade")
    if skills_col is None or grade_col is None:
        return
    for row in ws.iter_rows(min_row=2):
        row_num = row[0].row
        grade = ws.cell(row=row_num, column=grade_col).value
        expected = GRADE_TO_SKILLS.get(str(grade).strip().upper() if grade else None)
        if expected is None:
            continue
        skills_cell = ws.cell(row=row_num, column=skills_col)
        if skills_cell.value != expected:
            log.fixed(
                f"{ws.title}!{skills_cell.coordinate}: Skills '{skills_cell.value}' -> '{expected}' "
                f"(to match Grade {grade})"
            )
            skills_cell.value = expected


# --------------------------------------------------------------------------
# Candidate Tracker specific cleaning (dates, Joining Date, missing fields)
# --------------------------------------------------------------------------

def clean_candidate_tracker_dates(ws, headers, log: Log):
    col_index = {h: i + 1 for i, h in enumerate(headers) if h}
    date_col = col_index.get("Date")
    joining_col = col_index.get("Joining Date")
    accepted_col = col_index.get("Accepted")
    name_col = col_index.get("Name")
    if date_col is None:
        return

    years = []
    for row in ws.iter_rows(min_row=2):
        parsed = parse_date(ws.cell(row=row[0].row, column=date_col).value)
        if parsed is not None:
            years.append(parsed.year)
    reference_year = max(set(years), key=years.count) if years else None

    rows_to_drop = []
    for row in ws.iter_rows(min_row=2):
        row_num = row[0].row
        if all(c.value in (None, "") for c in row):
            rows_to_drop.append(row_num)
            continue
        if name_col and not ws.cell(row=row_num, column=name_col).value:
            rows_to_drop.append(row_num)
            log.dropped_count += 0  # counted via note, not the family-filter counter
            log.note(f"[dropped] {ws.title}!row {row_num}: no candidate name -- removed")
            continue

        date_cell = ws.cell(row=row_num, column=date_col)
        parsed = parse_date(date_cell.value)
        if parsed is not None:
            if date_cell.value != parsed:
                date_cell.value = parsed
                date_cell.number_format = "DD.MM.YYYY"
                log.fixed_count += 1
            if reference_year is not None and parsed.year != reference_year:
                fixed_date = parsed.replace(year=reference_year)
                log.fixed(
                    f"{ws.title}!{date_cell.coordinate}: year corrected {parsed.year} -> "
                    f"{reference_year} (day/month kept)"
                )
                date_cell.value = fixed_date
                date_cell.number_format = "DD.MM.YYYY"
        elif isinstance(date_cell.value, str) and date_cell.value.strip():
            log.warn(f"{ws.title}!{date_cell.coordinate}: '{date_cell.value}' is not a recognizable date")

        if joining_col is None:
            continue
        jcell = ws.cell(row=row_num, column=joining_col)
        jval = jcell.value
        jparsed = parse_date(jval)
        if jparsed is not None:
            if jval != jparsed:
                jcell.value = jparsed
                jcell.number_format = "DD.MM.YYYY"
                log.fixed_count += 1
            continue
        if not (isinstance(jval, str) and jval.strip()):
            continue
        cleaned = normalize_whitespace(jval)
        outcome = ws.cell(row=row_num, column=accepted_col).value if accepted_col else None

        if outcome in NON_HIRE_STATUSES:
            log.fixed(f"{ws.title}!{jcell.coordinate}: cleared '{cleaned}' (outcome is {outcome})")
            jcell.value = None
            continue

        month = parse_month_name(cleaned)
        row_date = ws.cell(row=row_num, column=date_col).value
        row_year = row_date.year if isinstance(row_date, datetime) else reference_year
        months_match = MONTHS_PATTERN.match(cleaned)

        if month is not None and row_year:
            new_value = f"{cleaned.strip().title()} {row_year}"
            log.fixed(f"{ws.title}!{jcell.coordinate}: '{cleaned}' -> '{new_value}' (year inferred)")
            jcell.value = new_value
        elif months_match and isinstance(row_date, datetime):
            n = int(months_match.group(1))
            derived = add_months(row_date, n)
            log.fixed(f"{ws.title}!{jcell.coordinate}: '{cleaned}' -> {derived.date().isoformat()} "
                      f"(interview date + {n} month(s), estimate)")
            jcell.value = derived
            jcell.number_format = "DD.MM.YYYY"
        else:
            jcell.value = cleaned
            log.warn(f"{ws.title}!{jcell.coordinate}: '{cleaned}' is not a recognizable date")

    for row_num in reversed(rows_to_drop):
        ws.delete_rows(row_num)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

def clean_workbook(input_path: Path, output_path: Path) -> Log:
    wb = openpyxl.load_workbook(input_path, data_only=False)
    log = Log()

    log.note(f"Sheets in source: {wb.sheetnames}")

    keep = {sheet_key(name) for name in SHEETS_TO_KEEP}
    to_remove = [name for name in wb.sheetnames if sheet_key(name) not in keep]
    for name in to_remove:
        log.note(f"[deleted]  {name} (not in SHEETS_TO_KEEP)")
        del wb[name]
    missing = keep - {sheet_key(name) for name in wb.sheetnames}
    if missing:
        log.warn(f"SHEETS_TO_KEEP names not found in this workbook (typo?): {sorted(missing)}")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row < 1:
            continue

        key = sheet_key(sheet_name)
        if key != sheet_name:
            log.fixed(f"Sheet name '{sheet_name}' -> '{key}' (stray whitespace in tab name)")
            ws.title = key

        headers = clean_headers(ws, log)

        for row in ws.iter_rows(min_row=2):
            for cell, header in zip(row, headers):
                if header in ("Role Applied", "Role", "Skills") and key in ROLE_ALIASES_BY_SHEET:
                    continue  # handled by filter_role_rows, not generic cleanup
                new_value = clean_cell_generic(cell, header, log)
                if new_value != cell.value:
                    cell.value = new_value

        clear_stray_punctuation_rows(ws, headers, log)
        filter_role_rows(ws, headers, log)

        if key == "Offering Sheet":
            fix_offering_sheet_skills(ws, headers, log)
        if key == CANDIDATE_SHEET:
            clean_candidate_tracker_dates(ws, headers, log)

        if ws.tables:
            resize_tables_to_fit(ws, log)

    wb.save(output_path)
    return log


def _resolve_paths():
    """Figure out input/output paths from CLI args or the hardcoded
    INPUT_PATH / OUTPUT_PATH config, with a clear message for every way
    this can be unconfigured or wrong -- raises SystemExit with the message
    already printed, never a bare traceback, for whichever is missing."""
    if len(sys.argv) >= 2:
        input_path = Path(sys.argv[1])
    else:
        if "PASTE_THE_FULL_PATH" in INPUT_PATH:
            print("No input file configured.\n")
            print("Open this script and edit the INPUT_PATH line near the top to point at")
            print("today's raw export -- look for '>>> EDIT THIS <<<' a few lines down from")
            print("the imports. Or pass a path directly:")
            print("\n    python extract_sd_analyst_german.py path\\to\\input.xlsx path\\to\\output.xlsx\n")
            raise SystemExit(1)
        input_path = Path(INPUT_PATH)

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        print("Check the INPUT_PATH line near the top of this script (or the path you passed in) "
              "-- it must be the FULL path, including the filename and .xlsx extension.")
        raise SystemExit(1)

    if len(sys.argv) >= 3:
        output_path = Path(sys.argv[2])
    else:
        if "PASTE_THE_FULL_PATH" in OUTPUT_PATH:
            print("No output path configured.\n")
            print("Open this script and edit the OUTPUT_PATH line near the top -- it must be a")
            print("full path ending in .xlsx, e.g.:")
            print(r'    OUTPUT_PATH = r"C:\Users\yourname\Documents\SD_Analyst_German_clean.xlsx"')
            raise SystemExit(1)
        output_path = Path(OUTPUT_PATH)

    if output_path.suffix.lower() != ".xlsx":
        print(f"OUTPUT_PATH doesn't end in .xlsx: {output_path}")
        print("Add the filename and extension, e.g. ...\\SD_Analyst_German_clean.xlsx")
        raise SystemExit(1)
    if not output_path.parent.exists():
        print(f"The folder for OUTPUT_PATH doesn't exist: {output_path.parent}")
        print("Create that folder first, or point OUTPUT_PATH at a folder that already exists.")
        raise SystemExit(1)
    if output_path.resolve() == input_path.resolve():
        print("OUTPUT_PATH is the same file as INPUT_PATH -- refusing to overwrite your raw export.")
        print("Point OUTPUT_PATH at a different filename.")
        raise SystemExit(1)

    return input_path, output_path


def main():
    input_path, output_path = _resolve_paths()
    log_path = output_path.with_name(output_path.stem + f"_log_{datetime.now():%Y%m%d_%H%M%S}.txt")

    log = clean_workbook(input_path, output_path)
    log_path.write_text(log.render())

    print(f"Output:  {output_path}")
    print(f"Log:     {log_path}")
    print(f"Cells normalized: {log.fixed_count}   Rows dropped: {log.dropped_count}   "
          f"Items needing review: {len(log.warnings)}")
    if log.warnings:
        print("\nNEEDS REVIEW:")
        for w in log.warnings:
            print(" -", w)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (0, None) and len(sys.argv) < 2:
            input("\nPress Enter to close...")  # keep a double-clicked window open
        raise
    except Exception:
        import traceback
        traceback.print_exc()
        print("\nSomething went wrong -- the error above is the actual cause.")
        if len(sys.argv) < 2:
            input("Press Enter to close...")  # keep a double-clicked window open
        sys.exit(1)
