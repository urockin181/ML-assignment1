#!/usr/bin/env python3
"""
Clean the Audi onboarding workbook (Attendance, PKI, H-O Schedule, Laptop
Tracker, ...) into a Power BI-ready star schema.

The source workbook is built for humans: several months stacked in one sheet,
dates as column headers, meaning carried by cell colours (first day, home
exception, site vs. home office), formulas for first/last name, the same
person spelled differently on every sheet, emails replaced by dates, etc.
Power BI needs the opposite: one flat table per subject, one header row, one
row per fact, typed columns and a key that joins the tables together.

Output workbook (every sheet is also a named Excel Table of the same name, so
Power BI's navigator lists it directly):

    Dim_Employee      one row per person, with EmployeeID (the join key),
                      position / start date / wave, and the onboarding status
                      from PKI, Laptop Tracker, Audi ID and Training Plan
    Dim_Date          calendar table covering every date in the data
    Fact_Attendance   one row per employee per day (unpivoted month grids),
                      colour legend decoded into flags (late, first day, ...)
    Fact_HO_Schedule  one row per employee per scheduled day, colour decoded
                      into Work Location (Site / Home / Home Exception)

Every fix / assumption the script makes (name matches, invalid emails, blank
PKI values, ...) is printed to the console, not written to the workbook.

PKI, Laptop Tracker, Audi ID and Training Plan hold one value per person, so
they are folded into Dim_Employee rather than exported as separate tables.
Training Room (room bookings) and Tools Tracker (no data) are not exported.

Usage
-----
    python clean_audi_onboarding.py
        Uses INPUT_PATH and OUTPUT_PATH set near the top of this file.

    python clean_audi_onboarding.py INPUT.xlsx OUTPUT.xlsx [--year 2026]
        Overrides both paths for this run.

--year is only used when a date in the source has no year (H-O Schedule
headers such as "Mon 21/9"); by default it is taken
from the Attendance sheet. The input file is never modified.

Requires: pandas, openpyxl
"""

from __future__ import annotations

import argparse
import colorsys
import re
import sys
from collections import Counter
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# >>> EDIT THESE TWO LINES <<<  Used when the script is run without arguments
# (e.g. by pressing Run in VS Code / IDLE). Keep the r before the quotes.
INPUT_PATH = r"C:\Onboarding\Audi.xlsx"
OUTPUT_PATH = r"C:\Onboarding\Audi_Clean.xlsx"

SHEET_ATTENDANCE = "Attendance"
SHEET_PKI = "PKI"
SHEET_LAPTOP = "Laptop Tracker"
SHEET_HO = "H-O Schedule"
SHEET_AUDI_ID = "Audi ID"
SHEET_TRAINING_PLAN = "Training Plan"

# Names that refer to the same person but can't be matched automatically
# (different surname on different sheets, no email given). Key: name as written
# in the source (case/spacing ignored). Value: that person's email.
# Add a line here whenever the console notes report a wrong/new match.
NAME_ALIASES = {
    "ammar yasser": "ammar.ali@capgemini.com",
    "basel mohamed": "bassel.shalaby@capgemini.com",
    "haneen bekhiet": "hanin.ahmed@capgemini.com",
    "amany hassan": "amany.abdelazim@capgemini.com",
    "amany": "amany.abdelazim@capgemini.com",
    "sarah abdelwahid": "sara.a.mohamed@capgemini.com",
}

# People to leave out of the clean file entirely (name as written, or email).
# Spacing, capitals and hyphens are ignored: "Ibrahim El Tohamy" matches too.
EXCLUDE_PEOPLE = [
    "Ibrahim Eltohamy",
    "ibrahim.eltohamy@capgemini.com",
]

# Rows in name-only sheets that are not a person.
NOT_A_PERSON = {"new comers", "newcomers"}

EMAIL_RE = re.compile(r"^[\w.+'-]+@[\w-]+(\.[\w-]+)+$")
SHIFT_RE = re.compile(r"(\d{1,2}):?(\d{2})\s*-\s*(\d{1,2}):?(\d{2})")
DAY_MONTH_RE = re.compile(r"(\d{1,2})\s*/\s*(\d{1,2})")

# Fallback hex colours, used only when a source cell has no fill and its value is
# not in the sheet's own legend. Otherwise the output copies the source colours.
STATUS_COLORS = {
    "Attended": "#92D050", "Attended Late": "#F4B183", "Off": "#D9D9D9",
    "Annual Leave": "#FFC000", "Sick Leave": "#FFFF00", "Emergency Leave": "#00B0F0",
    "No Call No Show": "#FF0000", "Home Office Issue": "#A6A6A6", "Other": "#FFFFFF",
}
STATUS_ORDER = {s: i for i, s in enumerate(STATUS_COLORS, 1)}
LOCATION_COLORS = {"Site": "#92D050", "Home": "#FFFFFF", "Home Exception": "#FFFF00"}


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

class DQLog:
    """Collects every change/assumption so the user can review it."""

    def __init__(self) -> None:
        self.rows: list[dict] = []

    def add(self, sheet: str, where: str, issue: str, value=None, action: str = "") -> None:
        self.rows.append({
            "Sheet": sheet, "Location": where, "Issue": issue,
            "Original Value": None if value is None else str(value), "Action Taken": action,
        })

    def frame(self) -> pd.DataFrame:
        cols = ["Sheet", "Location", "Issue", "Original Value", "Action Taken"]
        return pd.DataFrame(self.rows, columns=cols)


def clean_text(value) -> str | None:
    """Trim, collapse inner whitespace/newlines; empty -> None."""
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return None
    text = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()
    return text or None


def proper_name(text: str) -> str:
    """'ahmed aboelfadl' -> 'Ahmed Aboelfadl', 'magd-el-din' -> 'Magd-El-Din'."""
    return " ".join("-".join(p[:1].upper() + p[1:].lower() for p in w.split("-")) for w in text.split())


def letters(text: str) -> str:
    return re.sub(r"[^a-z]", "", text.lower())


def to_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def clean_email(value) -> str | None:
    text = clean_text(value)
    if not text:
        return None
    text = text.lower().removeprefix("mailto:")
    return text if EMAIL_RE.match(text) else None


def fill_key(cell) -> str:
    """Comparable key for a cell's background colour."""
    fill = cell.fill
    if fill is None or fill.fill_type is None:
        return "none"
    color = fill.fgColor
    if color.type == "rgb":
        rgb = str(color.rgb).upper()
        return "none" if rgb in ("00000000", "FFFFFFFF") else f"rgb:{rgb[-6:]}"
    if color.type == "theme":
        if color.theme == 0 and not color.tint:  # theme white == no colour
            return "none"
        return f"theme:{color.theme}:{round(color.tint or 0, 2)}"
    if color.type == "indexed":
        return "none" if color.indexed in (9, 64) else f"indexed:{color.indexed}"
    return "none"


THEME_SLOTS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4",
               "accent5", "accent6", "hlink", "folHlink"]


def theme_palette(wb) -> list[str]:
    """Workbook theme colours as RRGGBB, in the index order cell fills use."""
    xml = (wb.loaded_theme or b"").decode("utf-8", "ignore")
    palette = []
    for slot in THEME_SLOTS:
        match = re.search(rf"<a:{slot}>.*?(?:srgbClr val=\"(\w{{6}})\"|lastClr=\"(\w{{6}})\")", xml, re.S)
        palette.append((match.group(1) or match.group(2)).upper() if match else "FFFFFF")
    return palette


def apply_tint(rgb: str, tint: float) -> str:
    """Excel's tint: lighten (tint > 0) or darken (tint < 0) a colour's luminance."""
    r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "".join(f"{round(v * 255):02X}" for v in (r, g, b))


def fill_hex(cell, palette: list[str]) -> str | None:
    """A cell's background as '#RRGGBB', or None when it has no fill."""
    if fill_key(cell) == "none":
        return None
    color = cell.fill.fgColor
    if color.type == "rgb":
        return "#" + str(color.rgb)[-6:].upper()
    if color.type == "theme" and color.theme < len(palette):
        return "#" + apply_tint(palette[color.theme], color.tint or 0)
    return None


def yes_no(value) -> str:
    text = (clean_text(value) or "").lower()
    if text in ("yes", "y", "true", "1", "done", "received"):
        return "Yes"
    return "No"


def parse_shift(text: str) -> tuple[time | None, time | None, float | None]:
    match = SHIFT_RE.search(text or "")
    if not match:
        return None, None, None
    h1, m1, h2, m2 = map(int, match.groups())
    if h1 > 23 or h2 > 24 or m1 > 59 or m2 > 59:
        return None, None, None
    start, end = time(h1, m1), time(h2 % 24, m2)
    hours = ((h2 * 60 + m2) - (h1 * 60 + m1)) / 60
    if hours < 0:
        hours += 24
    return start, end, round(hours, 2)


# ---------------------------------------------------------------------------
# Employee registry: one EmployeeID per person, whatever sheet/spelling
# ---------------------------------------------------------------------------

class EmployeeRegistry:
    def __init__(self, log: DQLog) -> None:
        self.log = log
        self.people: list[dict] = []
        self.by_email: dict[str, int] = {}

    # -- signatures used to recognise the same person under a different spelling
    @staticmethod
    def name_signatures(name: str) -> set[str]:
        tokens = [letters(t) for t in re.split(r"[\s.\-_]+", name) if letters(t)]
        if not tokens:
            return set()
        sigs = {"".join(tokens), " ".join(sorted(tokens))}
        long_tokens = [t for t in tokens if len(t) > 1]  # drop middle initials
        if long_tokens:
            sigs |= {"".join(long_tokens), " ".join(sorted(long_tokens))}
        return sigs

    def signatures(self, idx: int) -> set[str]:
        person = self.people[idx]
        sigs: set[str] = set()
        for name in person["names"]:
            sigs |= self.name_signatures(name)
        if person["email"]:
            sigs |= self.name_signatures(person["email"].split("@")[0])
        return sigs

    def _first_names(self, idx: int) -> set[str]:
        return {letters(n.split()[0]) for n in self.people[idx]["names"] if n.split()}

    def _new(self, name: str, email: str | None) -> int:
        self.people.append({"names": [name], "email": email, "start": None})
        idx = len(self.people) - 1
        if email:
            self.by_email[email] = idx
        return idx

    def resolve(self, name: str | None, email: str | None, sheet: str, where: str,
                start: date | None = None, create: bool = True) -> int | None:
        """Return the person's index, creating a new person if unknown."""
        if not name and not email:
            return None
        idx = self._resolve(name, email, sheet, where, start)
        if idx is None:
            if not create:
                self.log.add(sheet, where, "Person not found in any tracker", name,
                             "Added as new employee (only appears in this sheet)")
            elif not email:
                self.log.add(sheet, where, "Person has no valid email", name,
                             "Added as new employee keyed on name")
            idx = self._new(name or email, email)
        person = self.people[idx]
        if name and name not in person["names"]:
            person["names"].append(name)
        if email and not person["email"]:
            person["email"] = email
            self.by_email[email] = idx
        return idx

    def _resolve(self, name, email, sheet, where, start) -> int | None:
        if email and email in self.by_email:
            idx = self.by_email[email]
            if name and not self.name_signatures(name) & self.signatures(idx):
                self.log.add(sheet, where, "Name differs from main record (same email)", name,
                             f"Matched by email to {self.people[idx]['names'][0]}")
            return idx
        if email:
            # Unseen email: the same person may already exist from a sheet that had no email
            # for them (e.g. new joiners in Attendance). Attach the email instead of duplicating.
            if name:
                sigs = self.name_signatures(name)
                hits = {i for i in range(len(self.people))
                        if not self.people[i]["email"] and sigs & self.signatures(i)}
                if len(hits) == 1:
                    return hits.pop()
            return None  # a genuinely new person
        key = (name or "").lower()
        alias_email = NAME_ALIASES.get(key)
        if alias_email and alias_email in self.by_email:
            idx = self.by_email[alias_email]
            self.log.add(sheet, where, "Name differs from main record", name,
                         f"Matched via NAME_ALIASES to {self.people[idx]['names'][0]} ({alias_email})")
            return idx
        sigs = self.name_signatures(name)
        hits = {i for i in range(len(self.people)) if sigs & self.signatures(i)}
        if len(hits) == 1:
            return hits.pop()
        if len(hits) > 1:
            self.log.add(sheet, where, "Name matches several people", name, "Kept as separate person")
            return None
        # Fallback: a first name that only ONE known person has (e.g. 'Sarah X' vs 'Sara Y'),
        # and the start date must agree if known. Common first names (Mohamed, Ahmed, ...)
        # shared by several people never match this way.
        first = letters(name.split()[0])
        hits = {i for i in range(len(self.people)) if first in self._first_names(i)}
        if len(hits) == 1:
            idx = hits.pop()
            if start and self.people[idx]["start"] and self.people[idx]["start"] != start:
                return None
            self.log.add(sheet, where, "Fuzzy match on first name - please verify", name,
                         f"Matched to {self.people[idx]['names'][0]}; add to NAME_ALIASES if correct")
            return idx
        return None


# ---------------------------------------------------------------------------
# Sheet parsers
# ---------------------------------------------------------------------------

def find_cell(ws, text: str):
    target = text.lower()
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().lower() == target:
                return cell
    return None


def header_map(ws, row: int) -> dict[str, int]:
    """'Got PKI' -> column index for a header row."""
    out = {}
    for cell in ws[row]:
        text = clean_text(cell.value)
        if text and text.lower() not in out:
            out[text.lower()] = cell.column
    return out


def attendance_legend(ws, palette: list[str]) -> tuple[dict[str, str], dict[str, str]]:
    """Read the 'Color Code' legend: fill colour -> meaning, and status -> its legend colour."""
    legend: dict[str, str] = {}
    status_hex: dict[str, str] = {}
    anchor = find_cell(ws, "Color Code")
    if anchor is None:
        return legend, status_hex
    r = anchor.row + 1
    while True:
        cell = ws.cell(r, anchor.column)
        label = clean_text(cell.value)
        if not label:
            break
        note = clean_text(ws.cell(r, anchor.column + 1).value)
        legend.setdefault(fill_key(cell), note or label)
        if fill_hex(cell, palette):
            status_hex.setdefault(normalize_status(label)["Status"], fill_hex(cell, palette))
        r += 1
    return legend, status_hex


def normalize_status(raw: str) -> dict:
    low = raw.lower()
    out = {"Status": "Other", "Status Category": "Other", "Is Present": 0, "Is Late": 0,
           "Is Tech Issue": int("tech issue" in low), "Is Home Office Issue": int("home office issue" in low)}
    if low.startswith("home office issue"):
        if "attend" in low:
            out.update({"Status": "Attended", "Status Category": "Present", "Is Present": 1})
        else:
            out.update({"Status": "Home Office Issue", "Status Category": "Unplanned Absence"})
    elif re.match(r"attended\s+l\b", low):
        out.update({"Status": "Attended Late", "Status Category": "Present", "Is Present": 1, "Is Late": 1})
    elif low.startswith("attend"):
        out.update({"Status": "Attended", "Status Category": "Present", "Is Present": 1})
    elif low in ("off", "day off", "weekend"):
        out.update({"Status": "Off", "Status Category": "Day Off"})
    elif low.startswith("annual"):
        out.update({"Status": "Annual Leave", "Status Category": "Planned Leave"})
    elif low.startswith("sick"):
        out.update({"Status": "Sick Leave", "Status Category": "Unplanned Absence"})
    elif low.startswith("emergency"):
        out.update({"Status": "Emergency Leave", "Status Category": "Unplanned Absence"})
    elif low in ("ncns", "no call no show"):
        out.update({"Status": "No Call No Show", "Status Category": "Unplanned Absence"})
    return out


def parse_attendance(wb, reg: EmployeeRegistry, log: DQLog) -> tuple[pd.DataFrame, dict]:
    ws = wb[SHEET_ATTENDANCE]
    palette = theme_palette(wb)
    legend, legend_hex = attendance_legend(ws, palette)
    blocks = [c.row for c in ws["A"] if clean_text(c.value) and c.value.strip().lower() == "name"]
    # Latest month first, so the newest spelling of a name becomes the display name.
    blocks.sort(reverse=True)
    info: dict[int, dict] = {}
    rows = []
    for header_row in blocks:
        hdr = header_map(ws, header_row)
        date_cols = {c.column: to_date(c.value) for c in ws[header_row] if to_date(c.value)}
        r = header_row + 1
        while True:
            name = clean_text(ws.cell(r, 1).value)
            if not name or name.lower() == "name":
                break
            where = f"row {r}"
            raw_email = ws.cell(r, hdr.get("email", 4)).value
            email = clean_email(raw_email)
            if raw_email is not None and not email:
                log.add(SHEET_ATTENDANCE, where, "Invalid email", raw_email, "Email left blank; matched by name")
            elif isinstance(raw_email, str) and raw_email != raw_email.strip():
                log.add(SHEET_ATTENDANCE, where, "Whitespace/newline in email", repr(raw_email), "Trimmed")
            start = to_date(ws.cell(r, hdr.get("start date", 5)).value)
            idx = reg.resolve(name, email, SHEET_ATTENDANCE, where, start)
            person = reg.people[idx]
            person["start"] = person["start"] or start
            meta = info.setdefault(idx, {})
            meta.setdefault("Position", clean_text(ws.cell(r, hdr.get("position", 6)).value))
            meta.setdefault("Start Date", start)
            for col, day in date_cols.items():
                cell = ws.cell(r, col)
                raw = clean_text(cell.value)
                if not raw:
                    continue  # before start date / not yet recorded
                status = normalize_status(raw)
                if status["Status"] == "Other":
                    log.add(SHEET_ATTENDANCE, cell.coordinate, "Unknown attendance value", raw, "Kept as 'Other'")
                note = legend.get(fill_key(cell))
                note_low = (note or "").lower()
                late_type = None
                if status["Is Late"] or "late" in note_low:
                    late_type = note if note and "min" in note_low else "Late"
                rows.append({
                    "EmployeeID": idx, "Date": day, **status,
                    "Late Type": late_type,
                    "Is First Day": int("first day" in note_low),
                    "Is Home Exception": int("home exception" in note_low),
                    "Is Working Day": int(status["Status Category"] != "Day Off"),
                    "Color Note": note,
                    # The cell's own colour; unfilled cells (e.g. 'Off') take the legend's colour.
                    "Status Color": fill_hex(cell, palette) or legend_hex.get(status["Status"])
                    or STATUS_COLORS.get(status["Status"], "#FFFFFF"),
                    "Status Order": STATUS_ORDER.get(status["Status"], 99),
                    "Original Value": raw,
                })
            r += 1
    df = pd.DataFrame(rows)
    before = len(df)
    df = df.drop_duplicates(["EmployeeID", "Date"], keep="first")
    if len(df) != before:
        log.add(SHEET_ATTENDANCE, "-", "Duplicate employee/date entries", before - len(df), "Kept first occurrence")
    return df, info


def parse_tracker(wb, sheet: str, value_header: str, reg: EmployeeRegistry, log: DQLog,
                  extra_headers: tuple[str, ...] = ()) -> pd.DataFrame:
    """PKI / Laptop Tracker: one row per person with a Yes/No column."""
    ws = wb[sheet]
    hdr = header_map(ws, 1)
    rows, seen = [], set()
    for r in range(2, ws.max_row + 1):
        name = clean_text(ws.cell(r, hdr.get("name", 1)).value)
        raw_email = ws.cell(r, hdr.get("email", 4)).value
        email = clean_email(raw_email)
        if not name:
            if raw_email is not None:
                log.add(sheet, f"row {r}", "Stray value outside the table", raw_email, "Ignored")
            continue
        if raw_email is not None and not email:
            log.add(sheet, f"row {r}", "Invalid email", raw_email, "Email left blank; matched by name")
        start = to_date(ws.cell(r, hdr.get("start date", 5)).value)
        idx = reg.resolve(name, email, sheet, f"row {r}", start)
        if idx in seen:
            log.add(sheet, f"row {r}", "Duplicate person", name, "Kept first occurrence")
            continue
        seen.add(idx)
        raw_value = ws.cell(r, hdr[value_header.lower()]).value
        if clean_text(raw_value) is None:
            log.add(sheet, f"row {r}", f"'{value_header}' is blank", name, "Treated as 'No'")
        row = {"EmployeeID": idx, "Name In Source": name, "Start Date In Source": start,
               value_header: yes_no(raw_value)}
        for h in extra_headers:
            row[h] = clean_text(ws.cell(r, hdr[h.lower()]).value) if h.lower() in hdr else None
        rows.append(row)
    return pd.DataFrame(rows)


def ho_legend(ws) -> dict[str, str]:
    """Legend cells like 'Green' | 'Site' -> fill colour -> location."""
    legend: dict[str, str] = {}
    for row in ws.iter_rows():
        for cell in row:
            text = clean_text(cell.value)
            if text and text.lower() in ("green", "white", "yellow", "red", "blue", "orange", "grey", "gray"):
                meaning = clean_text(ws.cell(cell.row, cell.column + 1).value)
                if meaning:
                    legend.setdefault(fill_key(cell), meaning)
    legend.setdefault("none", "Home")
    return legend


def header_date(value, year: int) -> date | None:
    if to_date(value):
        return to_date(value)
    match = DAY_MONTH_RE.search(str(value or ""))
    if match:
        day, month = map(int, match.groups())
        try:
            return date(year, month, day)
        except ValueError:
            return None
    return None


def parse_ho_schedule(wb, reg: EmployeeRegistry, log: DQLog, year: int) -> pd.DataFrame:
    ws = wb[SHEET_HO]
    palette = theme_palette(wb)
    legend = ho_legend(ws)
    header_rows = [c.row for c in ws["A"] if clean_text(c.value) and c.value.strip().lower() == "name"]
    rows = []
    for header_row in header_rows:
        hdr = header_map(ws, header_row)
        wave_col = hdr.get("wave", 2)
        task_col = hdr.get("task")
        day_cols = list(range(wave_col + 1, wave_col + 8))  # Mon..Sun
        week_cell = next((c for c in ws[header_row] if isinstance(c.value, (int, float))
                          and 1 <= c.value <= 53), None)
        # Monday of the block: from the ISO week number, otherwise from the first readable header.
        monday = None
        if week_cell is not None:
            monday = date.fromisocalendar(year, int(week_cell.value), 1)
        for i, col in enumerate(day_cols):
            parsed = header_date(ws.cell(header_row, col).value, year)
            if parsed:
                derived = parsed - timedelta(days=i)
                if monday and derived != monday:
                    log.add(SHEET_HO, f"row {header_row}", "Header date disagrees with week number",
                            ws.cell(header_row, col).value, f"Used week number (Monday {monday})")
                monday = monday or derived
                break
        if monday is None:
            log.add(SHEET_HO, f"row {header_row}", "Could not determine week dates", None, "Block skipped")
            continue
        for i, col in enumerate(day_cols):
            text = clean_text(ws.cell(header_row, col).value)
            if text and header_date(text, year) is None and not to_date(ws.cell(header_row, col).value):
                log.add(SHEET_HO, ws.cell(header_row, col).coordinate, "Unreadable day header", text,
                        f"Date taken from week position ({monday + timedelta(days=i)})")
        r = header_row + 1
        while r <= ws.max_row:
            name = clean_text(ws.cell(r, 1).value)
            if not name or name.lower() == "name":
                break
            idx = reg.resolve(name, None, SHEET_HO, f"row {r}", create=False)
            wave = ws.cell(r, wave_col).value
            task = clean_text(ws.cell(r, task_col).value) if task_col else None
            for i, col in enumerate(day_cols):
                cell = ws.cell(r, col)
                text = clean_text(cell.value)
                if not text:
                    continue
                start, end, hours = parse_shift(text)
                location = legend.get(fill_key(cell))
                if location is None:
                    log.add(SHEET_HO, cell.coordinate, "Unknown colour in schedule", fill_key(cell),
                            "Work Location set to 'Unknown'")
                    location = "Unknown"
                is_shift = start is not None
                rows.append({
                    "EmployeeID": idx, "Date": monday + timedelta(days=i),
                    "Week Number": monday.isocalendar()[1],
                    "Wave": int(wave) if isinstance(wave, (int, float)) else None,
                    "Schedule Type": "Shift" if is_shift else proper_name(text),
                    "Work Location": location if is_shift else None,
                    "Is On Site": int(is_shift and location.lower() == "site"),
                    "Is Home": int(is_shift and location.lower().startswith("home")),
                    "Shift Start": start, "Shift End": end, "Shift Hours": hours,
                    "Location Color": fill_hex(cell, palette) or (
                        LOCATION_COLORS.get(location, "#FFFFFF") if is_shift
                        else STATUS_COLORS.get(normalize_status(text)["Status"], "#FFFFFF")),
                    "Weekly Task": task, "Original Value": text,
                })
            r += 1
    df = pd.DataFrame(rows)
    before = len(df)
    df = df.drop_duplicates(["EmployeeID", "Date"], keep="first")
    if len(df) != before:
        log.add(SHEET_HO, "-", "Employee scheduled twice on same date", before - len(df), "Kept first occurrence")
    return df


def parse_audi_id(wb, reg: EmployeeRegistry, log: DQLog) -> dict[int, str]:
    ws = wb[SHEET_AUDI_ID]
    out = {}
    for r in range(2, ws.max_row + 1):
        uid = clean_text(ws.cell(r, 1).value)
        first, last = clean_text(ws.cell(r, 2).value), clean_text(ws.cell(r, 3).value)
        if not uid:
            continue
        name = proper_name(" ".join(p for p in (first, last) if p))
        idx = reg.resolve(name, None, SHEET_AUDI_ID, f"row {r}", create=False)
        if idx in out:
            log.add(SHEET_AUDI_ID, f"row {r}", "Second Audi ID for same person", uid, "Kept first ID")
            continue
        out[idx] = uid.upper()
    return out


def parse_training_plan(wb, reg: EmployeeRegistry, log: DQLog) -> pd.DataFrame:
    ws = wb[SHEET_TRAINING_PLAN]
    hdr = header_map(ws, 1)
    rows = []
    for r in range(2, ws.max_row + 1):
        name = clean_text(ws.cell(r, hdr.get("name", 1)).value)
        if not name:
            continue
        is_person = name.lower() not in NOT_A_PERSON
        idx = reg.resolve(name, None, SHEET_TRAINING_PLAN, f"row {r}", create=False) if is_person else None
        wave = ws.cell(r, hdr.get("wave", 2)).value
        rows.append({
            "EmployeeID": idx, "Name In Source": name,
            "Wave": int(wave) if isinstance(wave, (int, float)) else None,
            "Task": clean_text(ws.cell(r, hdr.get("task", 3)).value),
            "Owner": clean_text(ws.cell(r, hdr.get("owner", 4)).value),
            "Aufgaben (Tasks)": clean_text(ws.cell(r, hdr.get("aufgaben", 5)).value),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Building the model
# ---------------------------------------------------------------------------

def build_dim_employee(reg, att_info, pki, laptop, audi_ids, plan, ho, log) -> pd.DataFrame:
    pki_map = pki.set_index("EmployeeID")["PKI Received"].to_dict() if len(pki) else {}
    lap = laptop.set_index("EmployeeID") if len(laptop) else pd.DataFrame()
    plan_people = plan.dropna(subset=["EmployeeID"]).drop_duplicates("EmployeeID").set_index("EmployeeID")
    ho_wave = ho.dropna(subset=["Wave"]).sort_values("Date").groupby("EmployeeID")["Wave"].last() if len(ho) else {}
    rows = []
    for idx, person in enumerate(reg.people):
        name = proper_name(person["names"][0])
        meta = att_info.get(idx, {})
        start = meta.get("Start Date") or person["start"]
        if idx in lap.index and start and lap.at[idx, "Start Date In Source"] \
                and lap.at[idx, "Start Date In Source"] != start:
            log.add(SHEET_LAPTOP, name, "Start date differs from Attendance",
                    lap.at[idx, "Start Date In Source"], f"Used Attendance start date {start}")
        wave = plan_people.at[idx, "Wave"] if idx in plan_people.index else None
        if (wave is None or pd.isna(wave)) and idx in getattr(ho_wave, "index", []):
            wave = ho_wave[idx]
        pki_status = {"Yes": "Received", "No": "Not Received"}.get(pki_map.get(idx), "Not Tracked")
        laptop_status = ("Received" if lap.at[idx, "Laptop Received"] == "Yes" else "Not Received") \
            if idx in lap.index else "Not Tracked"
        has_audi = idx in audi_ids
        done = int(pki_status == "Received") + int(laptop_status == "Received") + int(has_audi)
        tokens = name.split()
        rows.append({
            "EmployeeID": idx, "Employee Name": name,
            "First Name": tokens[0], "Last Name": tokens[-1] if len(tokens) > 1 else None,
            "Email": person["email"],
            "Position": meta.get("Position"),
            "Start Date": start,
            "Start Month": start.strftime("%Y-%m") if start else None,
            "Wave": int(wave) if wave is not None and not pd.isna(wave) else None,
            "Audi User ID": audi_ids.get(idx),
            "Has Audi ID": "Yes" if has_audi else "No",
            "PKI Status": pki_status,
            "Laptop Status": laptop_status,
            "Training Mode": lap.at[idx, "Training Mode"] if idx in lap.index else None,
            "Training Task": plan_people.at[idx, "Task"] if idx in plan_people.index else None,
            "Training Owner": plan_people.at[idx, "Owner"] if idx in plan_people.index else None,
            "PKI Received Flag": int(pki_status == "Received"),
            "Laptop Received Flag": int(laptop_status == "Received"),
            "Audi ID Flag": int(has_audi),
            "Onboarding Steps Done": done,
            "Onboarding Status": "Complete" if done == 3 else ("In Progress" if done else "Not Started"),
            "In Attendance Sheet": "Yes" if idx in att_info else "No",
            "Name Variants": "; ".join(sorted({proper_name(n) for n in person["names"]} - {name})) or None,
        })
    dim = pd.DataFrame(rows)
    # Stable, readable keys ordered by start date then name: E001, E002, ...
    dim = dim.sort_values(["Start Date", "Employee Name"], na_position="last").reset_index(drop=True)
    dim["New ID"] = [f"E{i:03d}" for i in range(1, len(dim) + 1)]
    return dim


def build_dim_date(dates: list[date], working_days: list[date]) -> pd.DataFrame:
    dates = [d for d in dates if d]
    first = min(dates).replace(day=1)
    last_month = max(dates).replace(day=28) + timedelta(days=4)
    last = last_month - timedelta(days=last_month.day)
    days = pd.date_range(first, last, freq="D")
    iso = days.isocalendar()
    recent = set(sorted(set(working_days))[-10:])
    return pd.DataFrame({
        "Date": days.date,
        "Year": days.year,
        "Month Number": days.month,
        "Month": days.strftime("%B"),
        "Month Short": days.strftime("%b"),
        "Year Month": days.strftime("%Y-%m"),
        "ISO Week": iso.week.values,
        "Week Label": [f"W{w:02d}" for w in iso.week.values],
        "Week Start": (days - pd.to_timedelta(days.weekday, unit="D")).date,
        "Day": days.day,
        "Day Of Week Number": days.weekday + 1,
        "Day Name": days.strftime("%A"),
        "Day Short": days.strftime("%a"),
        "Is Weekend": [int(d >= 5) for d in days.weekday],
        # The 10 most recent days with attendance recorded, for the "last 10 working days" view.
        "Last 10 Working Days": ["Yes" if d in recent else "No" for d in days.date],
    })


def remap_ids(frames: dict[str, pd.DataFrame], mapping: dict[int, str]) -> None:
    for df in frames.values():
        if "EmployeeID" in df.columns:
            df["EmployeeID"] = df["EmployeeID"].map(lambda v: mapping.get(v) if pd.notna(v) else None)


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

# Output cells painted with the colour held in the same row's colour column,
# so the clean file shows the same colour code as the source sheets.
PAINTED_COLUMNS = {
    "Fact_Attendance": ("Status Color", ("Status", "Original Value", "Status Color")),
    "Fact_HO_Schedule": ("Location Color", ("Schedule Type", "Work Location", "Original Value", "Location Color")),
}


def paint_cells(ws, df: pd.DataFrame, color_col: str, targets: tuple[str, ...]) -> None:
    cols = {c: i for i, c in enumerate(df.columns, 1)}
    for row_idx, hex_color in enumerate(df[color_col], 2):
        if not isinstance(hex_color, str) or hex_color.upper() == "#FFFFFF":
            continue
        rgb = hex_color.lstrip("#").upper()
        fill = PatternFill(fill_type="solid", start_color="FF" + rgb, end_color="FF" + rgb)
        r, g, b = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
        dark = (0.299 * r + 0.587 * g + 0.114 * b) < 110
        for name in targets:
            if name in cols:
                cell = ws.cell(row_idx, cols[name])
                cell.fill = fill
                if dark:
                    cell.font = Font(color="FFFFFFFF")


def write_output(path: Path, tables: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in tables.items():
            df.to_excel(writer, sheet_name=name, index=False)
        wb = writer.book
        for name, df in tables.items():
            ws = wb[name]
            ws.freeze_panes = "A2"
            for col_idx, col in enumerate(df.columns, 1):
                letter = get_column_letter(col_idx)
                sample = df[col].dropna()
                fmt = None
                if len(sample) and isinstance(sample.iloc[0], date) and not isinstance(sample.iloc[0], datetime):
                    fmt = "yyyy-mm-dd"
                elif len(sample) and isinstance(sample.iloc[0], time):
                    fmt = "hh:mm"
                if fmt:
                    for cell in ws[letter][1:]:
                        cell.number_format = fmt
                width = max([len(str(col))] + [len(str(v)) for v in sample.head(200)]) + 2
                ws.column_dimensions[letter].width = min(max(width, 10), 60)
            if len(df):
                ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
                table = Table(displayName=name, ref=ref)
                table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
                ws.add_table(table)
            if name in PAINTED_COLUMNS and len(df):
                paint_cells(ws, df, *PAINTED_COLUMNS[name])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("input", type=Path, nargs="?", default=Path(INPUT_PATH),
                        help="path to the source workbook (default: INPUT_PATH in this file)")
    parser.add_argument("output", type=Path, nargs="?", default=Path(OUTPUT_PATH),
                        help="path for the cleaned workbook (default: OUTPUT_PATH in this file)")
    parser.add_argument("--year", type=int, help="year for dates written without one (default: from Attendance)")
    args = parser.parse_args(argv)

    if not args.input.exists():
        parser.error(f"input file not found: {args.input}")
    if args.input.resolve() == args.output.resolve():
        parser.error("output must be a different file than input")

    wb = load_workbook(args.input)  # formulas are recomputed here, cached values are not needed
    for required in (SHEET_ATTENDANCE, SHEET_PKI, SHEET_LAPTOP, SHEET_HO):
        if required not in wb.sheetnames:
            parser.error(f"sheet '{required}' not found; sheets are: {wb.sheetnames}")

    log = DQLog()
    reg = EmployeeRegistry(log)

    attendance, att_info = parse_attendance(wb, reg, log)
    year = args.year or Counter(d.year for d in attendance["Date"]).most_common(1)[0][0]
    pki = parse_tracker(wb, SHEET_PKI, "Got PKI", reg, log).rename(columns={"Got PKI": "PKI Received"})
    laptop = parse_tracker(wb, SHEET_LAPTOP, "Laptop received", reg, log, ("Attended Training",)).rename(
        columns={"Laptop received": "Laptop Received", "Attended Training": "Training Mode"})
    ho = parse_ho_schedule(wb, reg, log, year)
    audi_ids = parse_audi_id(wb, reg, log) if SHEET_AUDI_ID in wb.sheetnames else {}
    plan = parse_training_plan(wb, reg, log) if SHEET_TRAINING_PLAN in wb.sheetnames else pd.DataFrame(
        columns=["EmployeeID", "Wave", "Task", "Owner"])

    for sheet in wb.sheetnames:
        if sheet not in (SHEET_ATTENDANCE, SHEET_PKI, SHEET_LAPTOP, SHEET_HO, SHEET_AUDI_ID,
                         SHEET_TRAINING_PLAN):
            log.add(sheet, "-", "Sheet not needed for the dashboard", None, "Not exported")

    dim = build_dim_employee(reg, att_info, pki, laptop, audi_ids, plan, ho, log)
    excluded = {letters(x) for x in EXCLUDE_PEOPLE}
    drop = dim["Employee Name"].map(letters).isin(excluded) | dim["Email"].fillna("").map(letters).isin(excluded)
    for name in dim.loc[drop, "Employee Name"]:
        log.add("-", name, "Listed in EXCLUDE_PEOPLE", None, "Removed from every table")
    dropped_ids = set(dim.loc[drop, "EmployeeID"])
    dim = dim[~drop].reset_index(drop=True)
    dim["New ID"] = [f"E{i:03d}" for i in range(1, len(dim) + 1)]
    attendance = attendance[~attendance["EmployeeID"].isin(dropped_ids)].copy()
    ho = ho[~ho["EmployeeID"].isin(dropped_ids)].copy()
    mapping = dict(zip(dim["EmployeeID"], dim["New ID"]))
    dim = dim.drop(columns="EmployeeID").rename(columns={"New ID": "EmployeeID"})
    dim = dim[["EmployeeID"] + [c for c in dim.columns if c != "EmployeeID"]]

    for df in (dim, ho):
        df["Wave"] = df["Wave"].astype("Int64")
    remap_ids({"a": attendance, "h": ho}, mapping)
    name_of = dict(zip(dim["EmployeeID"], dim["Employee Name"]))
    for df in (attendance, ho):
        df.insert(1, "Employee Name", df["EmployeeID"].map(name_of))
        df.sort_values(["Date", "EmployeeID"], inplace=True)

    all_dates = list(attendance["Date"]) + list(ho["Date"]) + list(dim["Start Date"].dropna())
    tables = {
        "Dim_Employee": dim,
        "Dim_Date": build_dim_date(all_dates, list(attendance.loc[attendance["Is Working Day"] == 1, "Date"])),
        "Fact_Attendance": attendance,
        "Fact_HO_Schedule": ho,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    write_output(args.output, tables)

    print(f"Wrote {args.output}")
    for name, df in tables.items():
        print(f"  {name:<18} {len(df):>5} rows")
    print(f"\n{len(log.rows)} data-quality notes (check 'Fuzzy match' and 'Added as new employee'):")
    for note in log.rows:
        value = f" [{note['Original Value']}]" if note["Original Value"] else ""
        print(f"  {note['Sheet']} {note['Location']}: {note['Issue']}{value} -> {note['Action Taken']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
