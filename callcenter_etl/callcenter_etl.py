#!/usr/bin/env python3
"""
Call-center ETL for Power BI (Genesys Cloud CSV exports).

Processes four exports and writes ONE Excel workbook with four sheets:

    service_level_metrics_clean   <- queue performance export   (e.g. AudiQPer-Oct26.csv)
    abandoned_calls_clean         <- conversation detail export (e.g. AudiAban-Oct26.csv)
    agent_status_clean            <- agent status export        (e.g. AgentsSTA-26Oct26.csv)
    agent_performance_clean       <- agent performance export   (e.g. AgentsPer-26Oct26.csv)

Both agent sheets have exactly ONE row per agent per day: daily totals
(seconds and hh:mm:ss) and correctly weighted daily averages.

History / appending
-------------------
If the output workbook already exists, each newly processed dataset is merged
into the existing sheet using a "replacement key":

    service_level_metrics_clean : report_date + queue_name
    abandoned_calls_clean       : conversation_id
    agent_status_clean          : report_date + agent_id
    agent_performance_clean     : report_date + agent_id

Rows in the new file replace older rows that have the same key (a re-export
is treated as the newer, more correct version). All other historical rows are
kept, so daily agent files accumulate day by day and monthly queue files can
be re-run at any time without creating duplicates. One exception: a row from
an interval that is still in progress (Interval Complete = FALSE) never
replaces a row from a completed interval.

Input folders
-------------
Set the four input folders and the output folder in the block marked
">>> EDIT YOUR FOLDER PATHS HERE <<<" below. Each run processes every CSV in
those folders that has not been processed before (oldest first) and records it
in the workbook's run_log sheet. A file that is overwritten with new content is
processed again. An export type with no new files keeps its sheet unchanged.

Usage
-----
    python callcenter_etl.py                      # uses the folders set below
    python callcenter_etl.py --reprocess-all      # process every file again
    python callcenter_etl.py --files-only --agent-status AgentsSTA-26Oct26.csv   # one file

Run with --help for all options. Other defaults live in the Config class below.

Output
------
ONE Excel file and nothing else: <output folder>/callcenter_powerbi_clean.xlsx

    service_level_metrics_clean   \
    abandoned_calls_clean          |  data sheets (all records from every run)
    agent_status_clean             |
    agent_performance_clean       /
    run_log                       one row per file per run (also used to skip
                                  files that were already processed)
    data_quality_log              warnings / errors found in every run

Progress messages are printed to the console only.

Requires: Python 3.9+, pandas, openpyxl.
"""
from __future__ import annotations

import argparse
import csv
import io
import logging
import os
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

log = logging.getLogger("callcenter_etl")


# =============================================================================
# >>> EDIT YOUR FOLDER PATHS HERE <<<
#
# One folder per export type. Every day, drop the new CSV into its folder and
# run the script: it processes every CSV it has not processed before (oldest
# first) and adds it to the output workbook. Processed files are remembered in
# the workbook's run_log sheet, so nothing is loaded twice. The workbook is the
# only file the script writes.
#
# Use the full path inside r"..." (the r keeps Windows backslashes intact).
# Leave a path empty ("") to skip that export type.
# =============================================================================
SERVICE_LEVEL_FOLDER = r"C:\CallCenter\input\queue_performance"        # e.g. AudiQPer-*.csv
ABANDONED_FOLDER = r"C:\CallCenter\input\abandoned_calls"               # e.g. AudiAban-*.csv
AGENT_STATUS_FOLDER = r"C:\CallCenter\input\agent_status"               # e.g. AgentsSTA-*.csv
AGENT_PERFORMANCE_FOLDER = r"C:\CallCenter\input\agent_performance"     # e.g. AgentsPer-*.csv

OUTPUT_FOLDER = r"C:\CallCenter\output"   # callcenter_powerbi_clean.xlsx is written here


# =============================================================================
# Configuration
# =============================================================================
@dataclass
class Config:
    """All settings in one place. Command-line arguments override these."""

    # Input folders (set at the top of this file, or with --*-dir options).
    # Every CSV in a folder that was not processed before is processed.
    service_level_dir: str = SERVICE_LEVEL_FOLDER
    abandoned_dir: str = ABANDONED_FOLDER
    agent_status_dir: str = AGENT_STATUS_FOLDER
    agent_performance_dir: str = AGENT_PERFORMANCE_FOLDER

    # Optional single files (command line only); always processed when given.
    service_level_csv: Path | None = None
    abandoned_csv: Path | None = None
    agent_status_csv: Path | None = None
    agent_performance_csv: Path | None = None

    # True = ignore the run_log and process every file in the folders again.
    reprocess_all: bool = False

    output_dir: Path = Path(OUTPUT_FOLDER)
    output_workbook_name: str = "callcenter_powerbi_clean.xlsx"

    # Mandatory skill filter for the two agent files.
    target_skill: str = "GL AUDI EHD"
    # "token": keep agents whose ';'-separated skill list CONTAINS the skill
    #          (case/space-insensitive), e.g. "GL AUDI EHD; GL OTHER" is kept.
    # "exact": keep only rows whose whole Skills cell equals the skill.
    # Either way "GL AUDI VIP EHD" never matches "GL AUDI EHD".
    skill_match_mode: str = "token"

    # Date/time formats tried in order. The exports use e.g. "10/26/26 09:12 AM"
    # (month/day/2-digit-year). Put day-first formats first if that changes.
    datetime_formats: tuple[str, ...] = (
        "%m/%d/%y %I:%M %p",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%y %I:%M:%S %p",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%y %H:%M",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f",
    )
    # Encodings tried when the file has no byte-order mark.
    encodings: tuple[str, ...] = ("utf-8-sig", "cp1252", "latin-1")
    delimiters: tuple[str, ...] = (",", ";", "\t", "|")

    # Genesys leaves count cells blank when the value is zero. True = blank -> 0.
    blank_counts_as_zero: bool = True
    # Blank agent status cells mean "no time in this status". True = blank -> 0 s.
    blank_status_duration_as_zero: bool = True
    # Unit assumed when a duration cell is a bare number instead of HH:MM:SS.
    numeric_duration_unit: str = "seconds"  # "seconds" | "milliseconds" | "minutes"

    # Tolerances for consistency checks.
    duration_tolerance_seconds: float = 5.0
    pct_tolerance: float = 0.005  # 0.5 percentage points

    # Daily intervals: report_date = date of the interval midpoint. This puts
    # a DST-shifted day such as "10/25 11:00 PM -> 10/26 11:00 PM" on 26 Oct.
    # Shorter intervals use the date of Interval Start.
    daily_interval_min_hours: float = 20.0

    append_history: bool = True

    # Columns to drop from the output (e.g. ["remote"] to hide caller numbers).
    drop_columns: dict[str, list[str]] = field(default_factory=dict)


SHEETS = {
    "service_level": "service_level_metrics_clean",
    "abandoned": "abandoned_calls_clean",
    "agent_status": "agent_status_clean",
    "agent_performance": "agent_performance_clean",
}

# Replacement keys used when merging new rows into the existing workbook.
HISTORY_KEYS = {
    "service_level": ["report_date", "queue_name"],
    "abandoned": ["conversation_id"],
    "agent_status": ["report_date", "agent_id"],
    "agent_performance": ["report_date", "agent_id"],
}
COMPLETENESS_COLUMN = {
    "service_level": "interval_complete",
    "abandoned": None,
    "agent_status": "interval_complete",
    "agent_performance": "interval_complete",
}

# Column specs: canonical name -> (accepted normalized source names, required?)
# Source headers are normalized first ("Handled In SLA * " -> "handled_in_sla",
# "Answer %" -> "answer_pct", "Interval start" -> "interval_start").
SERVICE_LEVEL_SPEC = {
    "interval_start": (["interval_start"], True),
    "interval_end": (["interval_end"], False),
    "interval_complete": (["interval_complete"], False),
    "filters": (["filters"], False),
    "media_type": (["media_type"], False),
    "queue_id": (["queue_id", "queue_ids"], False),
    "queue_name": (["queue_name", "queue"], False),
    "offered": (["offer", "offered"], True),
    "answered": (["answer", "answered"], True),
    "answer_pct_source": (["answer_pct"], False),
    "abandoned": (["abandon", "abandoned"], True),
    "abandon_pct_source": (["abandon_pct"], False),
    "asa": (["asa"], False),
    "service_level_pct_source": (["service_level_pct"], True),
    "service_level_target_pct": (["service_level_target_pct"], False),
    "avg_wait": (["avg_wait"], False),
    "avg_handle": (["avg_handle"], False),
    "avg_talk": (["avg_talk"], False),
    "avg_hold": (["avg_hold"], False),
    "avg_acw": (["avg_acw"], False),
    "hold_count": (["hold"], False),
    "transfer_count": (["transfer"], False),
    "handled": (["handle", "handled"], False),
    "over_sla": (["over_sla"], False),
    "handled_in_sla": (["handled_in_sla"], False),
}

ABANDONED_SPEC = {
    "full_export_completed": (["full_export_completed"], False),
    "partial_result_timestamp": (["partial_result_timestamp"], False),
    "filters": (["filters"], False),
    "conversation_id": (["conversation_id"], True),
    "media_type": (["media_type"], False),
    "abandoned_raw": (["abandoned"], True),
    "time_to_abandon_raw": (["time_to_abandon"], True),
    "users": (["users", "agents"], False),
    "remote": (["remote", "ani"], False),
    "call_datetime_raw": (["date", "conversation_start", "start_time"], True),
    "duration_raw": (["duration"], False),
    "direction": (["direction"], False),
    "queue": (["queue", "queue_name"], False),
    "wrap_up": (["wrap_up"], False),
}

AGENT_STATUS_META_SPEC = {
    "interval_start": (["interval_start"], True),
    "interval_end": (["interval_end"], False),
    "interval_complete": (["interval_complete"], False),
    "filters": (["filters"], False),
    "agent_id": (["agent_id", "user_id"], False),
    "agent_name": (["agent_name", "user_name", "name"], False),
    "skills": (["skills", "skill"], True),
}

# Known status columns: normalized name -> (display name, status group, sort).
# Groups overlap by design, so only sum durations WITHIN a group:
#   Logged In = On Queue + Off Queue
#   Idle, Not Responding (and interacting time) happen while On Queue
#   Available, Away, Break, Meal, Busy, Meeting, Training are presence
#   statuses recorded while Off Queue
STATUS_DEFINITIONS = {
    "logged_in": ("Logged In", "1 Total Logged In", 1),
    "on_queue": ("On Queue", "2 Queue State", 2),
    "off_queue": ("Off Queue", "2 Queue State", 3),
    "idle": ("Idle", "3 On-Queue Routing Status", 4),
    "not_responding": ("Not Responding", "3 On-Queue Routing Status", 5),
    "interacting": ("Interacting", "3 On-Queue Routing Status", 6),
    "communicating": ("Communicating", "3 On-Queue Routing Status", 7),
    "available": ("Available", "4 Presence (Off Queue)", 8),
    "away": ("Away", "4 Presence (Off Queue)", 9),
    "break": ("Break", "4 Presence (Off Queue)", 10),
    "meal": ("Meal", "4 Presence (Off Queue)", 11),
    "busy": ("Busy", "4 Presence (Off Queue)", 12),
    "meeting": ("Meeting", "4 Presence (Off Queue)", 13),
    "training": ("Training", "4 Presence (Off Queue)", 14),
}
UNMAPPED_STATUS_GROUP = "9 Unmapped (review)"

NOT_RESPONDING_COUNT_NAMES = ["alert_no_answer", "alerts_no_answer", "alert_no_answer_count",
                              "not_responding", "not_responding_count", "no_answer"]

AGENT_PERFORMANCE_SPEC = {
    "interval_start": (["interval_start"], True),
    "interval_end": (["interval_end"], False),
    "interval_complete": (["interval_complete"], False),
    "filters": (["filters"], False),
    "media_type": (["media_type"], False),
    "agent_id": (["agent_id", "user_id"], False),
    "agent_name": (["agent_name", "user_name", "name"], False),
    "answered": (["answered", "answer"], True),
    "handled": (["handle", "handled"], True),
    "avg_handle_raw": (["avg_handle"], True),
    "avg_talk": (["avg_talk"], False),
    "avg_hold": (["avg_hold"], False),
    "avg_acw": (["avg_acw"], False),
    "held": (["held", "hold"], False),
    "transferred": (["transferred", "transfer"], False),
    # Number of Not Responding events. Not in the current export: add the Genesys column
    # "Alert - No Answer" to the Agent Performance view to fill it.
    "not_responding_count": (NOT_RESPONDING_COUNT_NAMES, False),
    "skills": (["skills", "skill"], True),
}

TRUE_VALUES = {"true", "yes", "y", "1", "t"}
FALSE_VALUES = {"false", "no", "n", "0", "f"}


class DatasetError(Exception):
    """A dataset cannot be processed (missing file, required column, ...)."""


# =============================================================================
# Data-quality report
# =============================================================================
class QualityReport:
    """Collects every check so it can be logged and written to a report."""

    LEVELS = {"OK": logging.INFO, "INFO": logging.INFO,
              "WARNING": logging.WARNING, "ERROR": logging.ERROR}

    def __init__(self) -> None:
        self.rows: list[dict] = []
        self.summary: dict[str, dict] = {}

    def add(self, dataset: str, check: str, status: str, count: int | None = None,
            detail: str = "", examples: str = "") -> None:
        self.rows.append({"dataset": dataset, "check": check, "status": status,
                          "affected_rows": count, "detail": detail, "examples": examples})
        msg = f"[{dataset}] {status}: {check}"
        if count is not None:
            msg += f" ({count} rows)"
        if detail:
            msg += f" - {detail}"
        log.log(self.LEVELS.get(status, logging.INFO), msg)

    def check(self, dataset: str, check: str, mask: pd.Series, df: pd.DataFrame,
              detail: str = "", severity: str = "WARNING",
              example_cols: list[str] | None = None) -> int:
        """Record a mask-based check. Returns the number of flagged rows."""
        mask = mask.fillna(False).astype(bool)
        n = int(mask.sum())
        if n == 0:
            self.add(dataset, check, "OK", 0)
            return 0
        examples = ""
        cols = [c for c in (example_cols or []) if c in df.columns]
        if "source_row" in df.columns:
            cols = ["source_row"] + cols
        if cols:
            sample = df.loc[mask, cols].head(5)
            examples = " | ".join(
                ", ".join(f"{c}={_fmt(v)}" for c, v in row.items())
                for _, row in sample.iterrows())
        self.add(dataset, check, severity, n, detail, examples)
        return n

    def set_summary(self, dataset: str, **values) -> None:
        self.summary.setdefault(dataset, {}).update(values)

    def to_frame(self) -> pd.DataFrame:
        return pd.DataFrame(self.rows, columns=["dataset", "check", "status",
                                                "affected_rows", "detail", "examples"])

    def log_rows(self, run_ts: pd.Timestamp) -> pd.DataFrame:
        """Non-OK checks of this run, for the data_quality_log sheet."""
        f = self.to_frame()
        f = f[f["status"] != "OK"].copy()
        f.insert(0, "run_timestamp", run_ts)
        return f


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NA or v is pd.NaT:
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d %H:%M") if (v.hour or v.minute) else v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        return f"{v:.4g}"
    return str(v)


# =============================================================================
# Reading and column handling
# =============================================================================
def detect_encoding(raw: bytes, candidates: tuple[str, ...]) -> str:
    """Use the byte-order mark when present, else the first encoding that decodes."""
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16"
    for enc in candidates:
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    raise DatasetError("Could not decode file with any of: " + ", ".join(candidates))


def detect_delimiter(header_line: str, candidates: tuple[str, ...]) -> str:
    """Pick the delimiter that splits the (quote-aware) header into most fields."""
    best, best_n = candidates[0], 0
    for d in candidates:
        n = len(next(csv.reader([header_line], delimiter=d)))
        if n > best_n:
            best, best_n = d, n
    if best_n < 2:
        raise DatasetError(f"Could not detect a delimiter from header: {header_line[:120]!r}")
    return best


def read_csv_robust(path: Path, cfg: Config, dq: QualityReport, dataset: str) -> pd.DataFrame:
    """
    Read a CSV as text (no type guessing), keeping a `source_row` line number.

    Rows with too few fields are padded and reported. Rows with extra non-empty
    fields cannot be aligned safely, so they are excluded and listed in the
    report. Fully blank lines are skipped.
    """
    path = Path(path)
    if not path.is_file():
        raise DatasetError(f"Input file not found: {path}")
    raw = path.read_bytes()
    if not raw.strip():
        raise DatasetError(f"Input file is empty: {path}")
    encoding = detect_encoding(raw, cfg.encodings)
    text = raw.decode(encoding)
    header_line = next((ln for ln in text.splitlines() if ln.strip()), "")
    delimiter = detect_delimiter(header_line, cfg.delimiters)

    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
    header: list[str] | None = None
    records, rows_idx, short_rows, malformed, blank = [], [], [], [], 0
    for row in reader:
        if not row or all(not c.strip() for c in row):
            blank += 1
            continue
        if header is None:
            header = row
            width = len(header)
            continue
        line = reader.line_num
        if len(row) < width:
            short_rows.append(line)
            row = row + [""] * (width - len(row))
        elif len(row) > width:
            if all(not c.strip() for c in row[width:]):
                row = row[:width]
            else:
                malformed.append(line)
                continue
        records.append(row)
        rows_idx.append(line)
    if header is None:
        raise DatasetError(f"No header row found in {path}")

    df = pd.DataFrame(records, columns=header, dtype=object)
    for c in df.columns:
        df[c] = df[c].map(lambda v: v.strip() if isinstance(v, str) else v)
    df = df.replace({"": None})
    df.insert(0, "source_row", rows_idx)

    log.info("[%s] Read %s: %d data rows, %d columns, encoding=%s, delimiter=%r",
             dataset, path.name, len(df), len(header), encoding, delimiter)
    if short_rows:
        dq.add(dataset, "Rows with fewer fields than the header (padded with blanks)", "WARNING",
               len(short_rows), examples="lines " + ", ".join(map(str, short_rows[:10])))
    if malformed:
        dq.add(dataset, "Malformed rows excluded (more fields than the header)", "ERROR",
               len(malformed), "These lines could not be aligned to columns; fix them in the source file.",
               "lines " + ", ".join(map(str, malformed[:10])))
    if not short_rows and not malformed:
        dq.add(dataset, "CSV structure (field counts match header)", "OK", 0)
    dq.set_summary(dataset, source_file=path.name, encoding=encoding,
                   rows_read=len(df), malformed_rows_excluded=len(malformed))
    return df


def normalize_name(name: str) -> str:
    """'Handled In SLA * ' -> 'handled_in_sla'; 'Answer %' -> 'answer_pct'."""
    s = str(name).replace("﻿", "").strip().lower()
    s = s.replace("%", " pct ").replace("#", " num ")
    s = re.sub(r"[^0-9a-z]+", "_", s)
    return s.strip("_")


def normalize_columns(df: pd.DataFrame, dataset: str) -> pd.DataFrame:
    """Normalize headers and stop if two different headers collide."""
    mapping = {c: normalize_name(c) if c != "source_row" else c for c in df.columns}
    seen: dict[str, str] = {}
    for orig, new in mapping.items():
        if new in seen:
            raise DatasetError(f"Columns {seen[new]!r} and {orig!r} both normalize to {new!r}; "
                               "rename one in the source or extend normalize_name().")
        if not new:
            raise DatasetError(f"Column {orig!r} has an empty name after normalization.")
        seen[new] = orig
    return df.rename(columns=mapping)


def resolve_columns(df: pd.DataFrame, spec: dict, dataset: str, dq: QualityReport,
                    expected_extra: set[str] | None = None) -> pd.DataFrame:
    """
    Rename source columns to canonical names using `spec`. Missing required
    columns raise DatasetError. Missing optional columns are added as empty
    columns so the output schema stays stable. Unmapped source columns are
    kept under their normalized names.
    """
    rename, missing_req, missing_opt = {}, [], []
    for canonical, (candidates, required) in spec.items():
        found = next((c for c in candidates if c in df.columns), None)
        if found is None:
            target = missing_req if required else missing_opt
            target.append(canonical if not required else f"{canonical} (source header like: {' / '.join(candidates)})")
        else:
            rename[found] = canonical
    if missing_req:
        raise DatasetError(
            f"Required column(s) missing: {missing_req}. Columns found: "
            f"{[c for c in df.columns if c != 'source_row']}")
    df = df.rename(columns=rename)
    for c in missing_opt:
        df[c] = None
    if missing_opt:
        dq.add(dataset, "Optional columns not found (added as empty)", "WARNING",
               detail=", ".join(missing_opt))
    extra = [c for c in df.columns
             if c not in spec and c != "source_row" and c not in (expected_extra or set())]
    if extra:
        dq.add(dataset, "Unrecognized source columns kept as-is", "INFO", detail=", ".join(extra))
    return df


# =============================================================================
# Value parsers (each returns parsed values plus masks of problems)
# =============================================================================
def parse_datetime(s: pd.Series, formats: tuple[str, ...]) -> tuple[pd.Series, pd.Series]:
    """Try each explicit format in order. Returns (datetimes, failed_mask)."""
    txt = s.astype(object).where(s.notna(), None)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for fmt in formats:
        todo = txt.notna() & out.isna()
        if not todo.any():
            break
        parsed = pd.to_datetime(txt[todo], format=fmt, errors="coerce")
        out.loc[parsed.index] = out.loc[parsed.index].fillna(parsed)
    failed = txt.notna() & out.isna()
    return out, failed


_DURATION_RE = re.compile(
    r"^(?:(?P<d>\d+)\s*d(?:ays?)?\s*)?(?:(?P<h>\d+):)?(?P<m>\d{1,2}):(?P<s>\d{1,2}(?:\.\d+)?)$",
    re.IGNORECASE)


def parse_duration_seconds(s: pd.Series, numeric_unit: str = "seconds") -> tuple[pd.Series, pd.Series]:
    """
    Convert durations to seconds (float). Accepts 'HH:MM:SS.fff', 'H:MM:SS',
    'MM:SS', '1d 02:03:04' and bare numbers (interpreted with `numeric_unit`).
    Returns (seconds, invalid_mask). Blank cells stay NaN and are not invalid.
    """
    txt = s.astype(object).where(s.notna(), None)
    txt = txt.map(lambda v: v.strip() if isinstance(v, str) else v)
    present = txt.notna()
    ext = txt[present].astype(str).str.extract(_DURATION_RE)
    secs = pd.Series(np.nan, index=s.index, dtype="float64")
    if len(ext):
        d = pd.to_numeric(ext["d"], errors="coerce").fillna(0)
        h = pd.to_numeric(ext["h"], errors="coerce").fillna(0)
        m = pd.to_numeric(ext["m"], errors="coerce")
        sec = pd.to_numeric(ext["s"], errors="coerce")
        bad_parts = (ext["h"].notna() & (m > 59)) | (sec >= 60)
        val = d * 86400 + h * 3600 + m * 60 + sec
        secs.loc[ext.index] = val.where(~bad_parts)
    # Bare numbers.
    unmatched = present & secs.isna()
    if unmatched.any():
        num = pd.to_numeric(txt[unmatched], errors="coerce")
        factor = {"seconds": 1, "milliseconds": 0.001, "minutes": 60}[numeric_unit]
        secs.loc[num.index] = num * factor
    invalid = present & (secs.isna() | (secs < 0))
    secs = secs.where(~invalid)
    return secs, invalid


def parse_count(s: pd.Series, blank_as_zero: bool) -> tuple[pd.Series, pd.Series]:
    """Non-negative whole numbers. Returns (Int64 series, invalid_mask)."""
    txt = s.astype(object).where(s.notna(), None)
    txt = txt.map(lambda v: v.replace(",", "").strip() if isinstance(v, str) else v)
    num = pd.to_numeric(txt, errors="coerce")
    invalid = txt.notna() & (num.isna() | (num < 0) | (num.round() != num))
    num = num.where(~invalid)
    if blank_as_zero:
        num = num.where(txt.notna(), 0)
    return num.astype("Float64").round().astype("Int64"), invalid


def parse_percent(s: pd.Series) -> tuple[pd.Series, pd.Series, bool]:
    """
    Return percentages as fractions (0.9333 = 93.33%). Values with a '%' sign
    are divided by 100. If a column has no '%' signs but contains values above
    1 (and none above 100) the whole column is taken as 0-100 and rescaled.
    Returns (fraction, invalid_mask, rescaled_flag).
    """
    txt = s.astype(object).where(s.notna(), None)
    txt = txt.map(lambda v: v.strip() if isinstance(v, str) else v)
    has_pct = txt.map(lambda v: isinstance(v, str) and v.endswith("%")).astype(bool)
    num = pd.to_numeric(txt.map(lambda v: v.rstrip("%").strip() if isinstance(v, str) else v),
                        errors="coerce").astype("float64")
    num = num.where(~has_pct, num / 100)
    plain = num[~has_pct].dropna()
    rescaled = bool(len(plain) and (plain > 1).any() and (plain <= 100).all())
    if rescaled:
        num = num.where(has_pct, num / 100)
    invalid = txt.notna() & (num.isna() | (num < 0) | (num > 1 + 1e-9))
    return num.where(~invalid), invalid, rescaled


def parse_bool(s: pd.Series) -> tuple[pd.Series, pd.Series]:
    """TRUE/FALSE, YES/NO, 1/0 ... -> nullable boolean. Returns (bool, unexpected_mask)."""
    norm = s.astype(object).map(lambda v: v.strip().lower() if isinstance(v, str) else v)
    out = pd.Series(pd.NA, index=s.index, dtype="boolean")
    out[norm.isin(TRUE_VALUES)] = True
    out[norm.isin(FALSE_VALUES)] = False
    unexpected = norm.notna() & out.isna()
    return out, unexpected


def skill_match(skills: pd.Series, target: str, mode: str) -> pd.Series:
    """True where the Skills cell matches the target skill (see Config.skill_match_mode)."""
    def norm(v: str) -> str:
        return re.sub(r"\s+", " ", v).strip().casefold()
    t = norm(target)
    if mode == "exact":
        return skills.map(lambda v: isinstance(v, str) and norm(v) == t).astype(bool)
    if mode == "token":
        return skills.map(lambda v: isinstance(v, str) and t in {norm(x) for x in v.split(";")}).astype(bool)
    raise ValueError(f"Unknown skill_match_mode: {mode!r}")


# =============================================================================
# Shared helpers
# =============================================================================
def build_flags(df: pd.DataFrame, checks: list[tuple[str, pd.Series]]) -> pd.Series:
    """Join the labels of every failed row-level check into one 'dq_flags' text."""
    if not checks:
        return pd.Series("", index=df.index)
    parts = [np.where(m.reindex(df.index).fillna(False).astype(bool).to_numpy(), label, "")
             for label, m in checks]
    return pd.Series(["; ".join(p for p in row if p) for row in zip(*parts)], index=df.index)


def add_interval_dimensions(df: pd.DataFrame, cfg: Config, dataset: str,
                            dq: QualityReport, flags: list) -> pd.DataFrame:
    """Parse interval start/end and add report_date, weekday, month, hour columns."""
    df["interval_start"], bad_start = parse_datetime(df["interval_start"], cfg.datetime_formats)
    df["interval_end"], bad_end = parse_datetime(df["interval_end"], cfg.datetime_formats)
    dq.check(dataset, "Unparseable Interval Start", bad_start, df, severity="ERROR")
    dq.check(dataset, "Unparseable Interval End", bad_end, df)
    dq.check(dataset, "Missing Interval Start", df["interval_start"].isna() & ~bad_start, df, severity="ERROR")
    flags += [("bad_interval_start", bad_start | df["interval_start"].isna()), ("bad_interval_end", bad_end)]

    length = df["interval_end"] - df["interval_start"]
    df["interval_minutes"] = (length.dt.total_seconds() / 60).round(2)
    neg = length < pd.Timedelta(0)
    dq.check(dataset, "Interval End before Interval Start", neg, df, severity="ERROR",
             example_cols=["interval_start", "interval_end"])
    flags.append(("interval_end_before_start", neg))

    daily = length >= pd.Timedelta(hours=cfg.daily_interval_min_hours)
    anchor = df["interval_start"].where(~daily, df["interval_start"] + length / 2)
    df["report_date"] = anchor.dt.normalize()
    shifted = df["report_date"].notna() & (df["report_date"] != df["interval_start"].dt.normalize())
    if shifted.any():
        dq.check(dataset, "Daily intervals not starting at midnight (report_date taken from interval midpoint)",
                 shifted, df, severity="INFO",
                 detail="Typical after a daylight-saving change: '10/25 11:00 PM -> 10/26 11:00 PM' "
                        "is reported as 26 Oct.",
                 example_cols=["interval_start", "interval_end", "report_date"])
    lengths = df.loc[df["interval_minutes"].notna(), "interval_minutes"].round().value_counts()
    dq.add(dataset, "Interval lengths found (minutes)", "INFO",
           detail=", ".join(f"{int(k)} min x{v}" for k, v in lengths.items()))
    add_date_parts(df, "report_date")
    df["interval_start_hour"] = df["interval_start"].dt.hour.astype("Int64")
    return df


def add_date_parts(df: pd.DataFrame, date_col: str) -> None:
    df["report_month"] = df[date_col].dt.strftime("%Y-%m")
    df["weekday"] = df[date_col].dt.day_name()


def apply_skill_filter(df: pd.DataFrame, cfg: Config, dataset: str, dq: QualityReport) -> pd.DataFrame:
    """Keep only the target skill and report exactly what was excluded."""
    before = len(df)
    keep = skill_match(df["skills"], cfg.target_skill, cfg.skill_match_mode)
    excluded = df[~keep]
    blank_skill = int(excluded["skills"].isna().sum())
    near = excluded["skills"].dropna()
    near = near[near.str.contains("audi", case=False)]
    dq.add(dataset, f"Skill filter '{cfg.target_skill}' (mode={cfg.skill_match_mode})", "INFO",
           detail=f"{before} rows before, {int(keep.sum())} kept, {before - int(keep.sum())} excluded "
                  f"({blank_skill} with blank Skills)")
    if len(near):
        dq.add(dataset, "Excluded rows whose Skills mention 'AUDI' (check they should be excluded)",
               "WARNING", len(near), detail="; ".join(sorted(near.unique())[:10]))
    if keep.sum() == 0:
        dq.add(dataset, "No rows left after skill filter", "ERROR", 0,
               detail="Check Config.target_skill / skill_match_mode against the Skills column.")
    dq.set_summary(dataset, rows_before_skill_filter=before, rows_after_skill_filter=int(keep.sum()))
    return df[keep].copy()


def check_agent_identity(df: pd.DataFrame, dataset: str, dq: QualityReport) -> None:
    """Agent id <-> name must be one-to-one."""
    if df["agent_id"].isna().all():
        dq.add(dataset, "Agent Id column empty - agent_name used as identifier", "WARNING")
        df["agent_id"] = df["agent_name"]
    dq.check(dataset, "Missing agent id", df["agent_id"].isna(), df, severity="ERROR",
             example_cols=["agent_name"])
    dq.check(dataset, "Missing agent name", df["agent_name"].isna(), df, example_cols=["agent_id"])
    pairs = df[["agent_id", "agent_name"]].dropna().drop_duplicates()
    multi_name = pairs["agent_id"].duplicated(keep=False)
    multi_id = pairs["agent_name"].duplicated(keep=False)
    if multi_name.any():
        dq.add(dataset, "Agent id with more than one name", "WARNING", int(multi_name.sum()),
               examples="; ".join(pairs.loc[multi_name, "agent_name"].head(5)))
    if multi_id.any():
        dq.add(dataset, "Agent name shared by different agent ids", "WARNING", int(multi_id.sum()),
               examples="; ".join(pairs.loc[multi_id, "agent_name"].head(5)))
    if not multi_name.any() and not multi_id.any():
        dq.add(dataset, "Agent id <-> name one-to-one", "OK", 0)


def check_duplicates(df: pd.DataFrame, key: list[str], dataset: str, dq: QualityReport,
                     flags: list) -> None:
    """Flag (never drop) exact duplicate rows and rows repeating the natural key."""
    content = [c for c in df.columns if c != "source_row"]
    exact = df.duplicated(subset=content, keep=False)
    by_key = df.duplicated(subset=key, keep=False) & ~exact
    dq.check(dataset, "Exact duplicate rows (flagged, not removed)", exact, df)
    dq.check(dataset, f"Rows sharing the key {key} with different values (flagged, not removed)",
             by_key, df, example_cols=key)
    flags += [("exact_duplicate", exact), ("duplicate_key_conflict", by_key)]


def add_source_columns(df: pd.DataFrame, path: Path, load_ts: pd.Timestamp) -> pd.DataFrame:
    df["source_file"] = Path(path).name
    df["load_timestamp"] = load_ts
    return df


def order_columns(df: pd.DataFrame, first: list[str]) -> pd.DataFrame:
    first = [c for c in first if c in df.columns]
    tail = [c for c in df.columns if c not in first and c not in
            ("dq_flags", "source_file", "source_row", "load_timestamp")]
    end = [c for c in ("dq_flags", "source_file", "source_row", "load_timestamp") if c in df.columns]
    return df[first + tail + end]


# =============================================================================
# 1. Service level / queue performance
# =============================================================================
def process_service_level(path: Path, cfg: Config, dq: QualityReport, load_ts) -> pd.DataFrame:
    ds = "service_level"
    df = normalize_columns(read_csv_robust(path, cfg, dq, ds), ds)
    df = resolve_columns(df, SERVICE_LEVEL_SPEC, ds, dq)
    flags: list[tuple[str, pd.Series]] = []

    df = add_interval_dimensions(df, cfg, ds, dq, flags)
    df["interval_complete"], bad = parse_bool(df["interval_complete"])
    dq.check(ds, "Unexpected Interval Complete values", bad, df)

    for c in ["offered", "answered", "abandoned", "handled", "over_sla", "handled_in_sla",
              "hold_count", "transfer_count"]:
        df[c], bad = parse_count(df[c], cfg.blank_counts_as_zero)
        dq.check(ds, f"Invalid count in '{c}' (non-numeric, negative or decimal)", bad, df)
        flags.append((f"invalid_{c}", bad))

    for c in ["answer_pct_source", "abandon_pct_source", "service_level_pct_source",
              "service_level_target_pct"]:
        df[c], bad, rescaled = parse_percent(df[c])
        if rescaled:
            dq.add(ds, f"'{c}' looked like 0-100 and was rescaled to a 0-1 fraction", "INFO")
        dq.check(ds, f"Invalid percentage in '{c}'", bad, df)
        flags.append((f"invalid_{c}", bad))

    for c in ["asa", "avg_wait", "avg_handle", "avg_talk", "avg_hold", "avg_acw"]:
        secs, bad = parse_duration_seconds(df[c], cfg.numeric_duration_unit)
        df[f"{c}_seconds"] = secs.round(3)
        dq.check(ds, f"Invalid duration in '{c}'", bad, df, example_cols=[c])
        flags.append((f"invalid_{c}", bad))
    df = df.drop(columns=["asa", "avg_wait", "avg_handle", "avg_talk", "avg_hold", "avg_acw"])

    # --- Derived metrics (definitions verified against the source columns) ---
    # Answered within SLA = Answered - Over SLA. This reproduces the source
    # 'Service Level %' exactly (19/28 = 67.86%, 4/7 = 57.14%). The source
    # column 'Handled In SLA *' only counts COMPLETED (handled) calls, so it
    # differs while calls are still active. It is kept as handled_in_sla.
    ok = df["over_sla"].notna() & df["answered"].notna() & (df["over_sla"] <= df["answered"])
    df["answered_within_sla"] = (df["answered"] - df["over_sla"]).where(ok).astype("Int64")
    ans = df["answered"].astype("Float64")
    off = df["offered"].astype("Float64")
    df["service_level_pct_calc"] = (df["answered_within_sla"] / ans.where(ans > 0)).astype("float64").round(6)
    df["answer_pct_calc"] = (ans / off.where(off > 0)).astype("float64").round(6)
    df["abandon_pct_calc"] = (df["abandoned"] / off.where(off > 0)).astype("float64").round(6)
    # Totals that make weighted averages possible in Power BI.
    df["total_handle_seconds"] = (df["avg_handle_seconds"] * df["handled"].astype("float64")).round(3)
    df["total_answer_wait_seconds"] = (df["asa_seconds"] * ans.astype("float64")).round(3)
    df["has_activity"] = df["offered"].fillna(0) > 0
    df["data_status"] = np.select(
        [df["interval_complete"].fillna(False).astype(bool),
         df["has_activity"]],
        ["Complete", "Incomplete interval (partial data)"],
        "Incomplete interval (no activity yet)")

    # --- Consistency checks ---
    tol = cfg.pct_tolerance
    over = (df["answered"] + df["abandoned"]) > df["offered"]
    dq.check(ds, "Answered + Abandoned > Offered", over, df,
             example_cols=["report_date", "offered", "answered", "abandoned"])
    gap = (df["offered"] - df["answered"] - df["abandoned"]) > 0
    dq.check(ds, "Offered > Answered + Abandoned (calls still waiting/flowed out)", gap, df,
             severity="INFO", example_cols=["report_date", "offered", "answered", "abandoned", "interval_complete"])
    bad_over = df["over_sla"] > df["answered"]
    dq.check(ds, "Over SLA > Answered", bad_over, df, example_cols=["report_date", "answered", "over_sla"])
    bad_hin = df["handled_in_sla"] > df["handled"]
    dq.check(ds, "Handled In SLA > Handled", bad_hin, df,
             example_cols=["report_date", "handled", "handled_in_sla"])
    sl_mis = (df["service_level_pct_calc"] - df["service_level_pct_source"]).abs() > tol
    dq.check(ds, "Service Level % does not equal (Answered - Over SLA) / Answered", sl_mis, df,
             example_cols=["report_date", "service_level_pct_source", "service_level_pct_calc"])
    ans_mis = (df["answer_pct_calc"] - df["answer_pct_source"]).abs() > tol
    dq.check(ds, "Answer % does not equal Answered / Offered", ans_mis, df,
             example_cols=["report_date", "answer_pct_source", "answer_pct_calc"])
    ab_src = df["abandon_pct_source"].fillna(0).where(df["offered"] > 0)
    ab_mis = (df["abandon_pct_calc"] - ab_src).abs() > tol
    dq.check(ds, "Abandon % does not equal Abandoned / Offered", ab_mis, df,
             example_cols=["report_date", "abandon_pct_source", "abandon_pct_calc"])
    hin_diff = df["handled_in_sla"].notna() & df["answered_within_sla"].notna() & \
        (df["handled_in_sla"] != df["answered_within_sla"])
    dq.check(ds, "'Handled In SLA *' differs from Answered - Over SLA", hin_diff, df, severity="INFO",
             detail="Use answered_within_sla for Service Level; handled_in_sla is kept as the source value.",
             example_cols=["report_date", "answered", "over_sla", "handled", "handled_in_sla", "interval_complete"])
    sl_missing = (df["answered"] > 0) & df["service_level_pct_source"].isna()
    dq.check(ds, "Answered > 0 but Service Level % blank", sl_missing, df)
    dq.check(ds, "Intervals not complete (partial or future data)",
             ~df["interval_complete"].fillna(False).astype(bool), df, severity="INFO",
             example_cols=["report_date", "offered"])
    target_blank = df["service_level_target_pct"].isna().all()
    if target_blank:
        dq.add(ds, "Service Level Target % is blank in every row", "INFO",
               detail="The SLA threshold (seconds) and target % are not in the file.")
    if df["queue_name"].notna().any():
        dq.add(ds, "Queues included in this export", "INFO",
               detail=" | ".join(df["queue_name"].dropna().unique()))
    flags += [("answered_plus_abandoned_gt_offered", over), ("over_sla_gt_answered", bad_over),
              ("sl_pct_mismatch", sl_mis), ("answer_pct_mismatch", ans_mis),
              ("abandon_pct_mismatch", ab_mis), ("handled_in_sla_gt_handled", bad_hin)]
    check_duplicates(df, ["interval_start", "interval_end", "queue_id"], ds, dq, flags)

    df["dq_flags"] = build_flags(df, flags)
    df = add_source_columns(df, path, load_ts)
    dq.set_summary(ds, output_rows=len(df), date_from=df["report_date"].min(),
                   date_to=df["report_date"].max())
    return order_columns(df, [
        "report_date", "report_month", "weekday", "interval_start", "interval_end",
        "interval_minutes", "interval_start_hour", "interval_complete", "data_status", "has_activity",
        "queue_name", "queue_id", "media_type",
        "offered", "answered", "abandoned", "answered_within_sla", "over_sla",
        "handled", "handled_in_sla", "hold_count", "transfer_count",
        "service_level_pct_source", "service_level_pct_calc", "service_level_target_pct",
        "answer_pct_source", "answer_pct_calc", "abandon_pct_source", "abandon_pct_calc",
        "asa_seconds", "total_answer_wait_seconds", "avg_wait_seconds",
        "avg_handle_seconds", "total_handle_seconds", "avg_talk_seconds",
        "avg_hold_seconds", "avg_acw_seconds", "filters"])


# =============================================================================
# 2. Abandoned calls (conversation detail)
# =============================================================================
def process_abandoned(path: Path, cfg: Config, dq: QualityReport, load_ts) -> pd.DataFrame:
    ds = "abandoned"
    df = normalize_columns(read_csv_robust(path, cfg, dq, ds), ds)
    df = resolve_columns(df, ABANDONED_SPEC, ds, dq)
    flags: list[tuple[str, pd.Series]] = []

    fec, bad = parse_bool(df["full_export_completed"])
    df["full_export_completed"] = fec
    partial = fec.eq(False).fillna(False).astype(bool) | df["partial_result_timestamp"].notna()
    dq.check(ds, "Export marked as partial (Full Export Completed = NO)", partial, df, severity="ERROR",
             detail="The export did not finish; rows may be missing.")

    df["call_datetime"], bad_dt = parse_datetime(df["call_datetime_raw"], cfg.datetime_formats)
    dq.check(ds, "Unparseable call date/time", bad_dt, df, severity="ERROR", example_cols=["call_datetime_raw"])
    dq.check(ds, "Missing call date/time", df["call_datetime_raw"].isna(), df, severity="ERROR",
             example_cols=["conversation_id"])
    df["report_date"] = df["call_datetime"].dt.normalize()
    add_date_parts(df, "report_date")
    df["call_hour"] = df["call_datetime"].dt.hour.astype("Int64")
    df["half_hour_slot"] = df["call_datetime"].dt.floor("30min").dt.strftime("%H:%M")

    df["is_abandoned"], bad_ab = parse_bool(df["abandoned_raw"])
    dq.check(ds, "Unexpected 'Abandoned' values (left blank, not guessed)", bad_ab | df["abandoned_raw"].isna(),
             df, severity="ERROR", example_cols=["abandoned_raw"])
    df["abandoned_flag"] = df["is_abandoned"].astype("Int64")  # 1/0 for SUM in Power BI

    df["time_to_abandon_seconds"], bad_tta = parse_duration_seconds(df["time_to_abandon_raw"],
                                                                    cfg.numeric_duration_unit)
    df["time_to_abandon_seconds"] = df["time_to_abandon_seconds"].round(3)
    df["duration_seconds"], bad_dur = parse_duration_seconds(df["duration_raw"], cfg.numeric_duration_unit)
    df["duration_seconds"] = df["duration_seconds"].round(3)
    dq.check(ds, "Invalid Time to Abandon", bad_tta, df, severity="ERROR", example_cols=["time_to_abandon_raw"])
    dq.check(ds, "Invalid Duration", bad_dur, df, example_cols=["duration_raw"])

    is_ab = df["is_abandoned"].fillna(False).astype(bool)
    not_ab = df["is_abandoned"].eq(False).fillna(False).astype(bool)
    ab_no_tta = is_ab & df["time_to_abandon_seconds"].isna()
    tta_not_ab = not_ab & df["time_to_abandon_seconds"].notna()
    tta_gt_dur = df["time_to_abandon_seconds"] > df["duration_seconds"] + cfg.duration_tolerance_seconds
    no_duration = not_ab & df["duration_seconds"].isna()
    dq.check(ds, "Abandoned call without Time to Abandon", ab_no_tta, df, example_cols=["conversation_id"])
    dq.check(ds, "Time to Abandon present on a non-abandoned call", tta_not_ab, df,
             example_cols=["conversation_id", "time_to_abandon_raw"])
    dq.check(ds, "Time to Abandon longer than call Duration", tta_gt_dur, df,
             example_cols=["conversation_id", "time_to_abandon_raw", "duration_raw"])
    dq.check(ds, "Non-abandoned call with blank Duration (probably still in progress at export time)",
             no_duration, df, example_cols=["conversation_id", "call_datetime", "users"])
    ab_with_agent = is_ab & df["users"].notna()
    dq.check(ds, "Abandoned call that has an agent in 'Users' (abandoned while alerting/after transfer)",
             ab_with_agent, df, severity="INFO", example_cols=["conversation_id", "users"])

    df["user_count"] = df["users"].map(
        lambda v: len([u for u in v.split(";") if u.strip()]) if isinstance(v, str) else 0).astype("Int64")
    df["first_user"] = df["users"].map(lambda v: v.split(";")[0].strip() if isinstance(v, str) else None)
    df["call_status"] = np.select(
        [is_ab, not_ab & df["duration_seconds"].isna(), not_ab],
        ["Abandoned", "Not abandoned - no duration (in progress?)", "Not abandoned"], "Unknown")

    for c, label in (("direction", "Directions"), ("queue", "Queues"), ("media_type", "Media types")):
        vals = df[c].dropna().unique()
        if len(vals):
            dq.add(ds, f"{label} present", "INFO", detail=", ".join(map(str, vals)))
    dq.check(ds, "Missing Conversation ID", df["conversation_id"].isna(), df, severity="ERROR")
    flags += [("bad_call_datetime", bad_dt | df["call_datetime_raw"].isna()),
              ("unexpected_abandoned_value", bad_ab | df["abandoned_raw"].isna()),
              ("invalid_time_to_abandon", bad_tta), ("invalid_duration", bad_dur),
              ("abandoned_without_time_to_abandon", ab_no_tta),
              ("time_to_abandon_on_non_abandoned", tta_not_ab),
              ("time_to_abandon_gt_duration", tta_gt_dur), ("no_duration", no_duration),
              ("partial_export", partial)]
    check_duplicates(df, ["conversation_id"], ds, dq, flags)

    df["dq_flags"] = build_flags(df, flags)
    df = add_source_columns(df, path, load_ts)
    df = df.drop(columns=["call_datetime_raw", "duration_raw"])
    n_ab = int(is_ab.sum())
    dq.set_summary(ds, output_rows=len(df), abandoned_calls=n_ab,
                   date_from=df["report_date"].min(), date_to=df["report_date"].max())
    return order_columns(df, [
        "conversation_id", "call_datetime", "report_date", "report_month", "weekday",
        "call_hour", "half_hour_slot", "is_abandoned", "abandoned_flag", "abandoned_raw",
        "call_status", "time_to_abandon_seconds", "time_to_abandon_raw", "duration_seconds",
        "queue", "direction", "media_type", "users", "user_count", "first_user", "remote",
        "wrap_up", "full_export_completed", "partial_result_timestamp", "filters"])


# =============================================================================
# 3. Agent status
# =============================================================================
def process_agent_status(path: Path, cfg: Config, dq: QualityReport, load_ts) -> pd.DataFrame:
    ds = "agent_status"
    df = normalize_columns(read_csv_robust(path, cfg, dq, ds), ds)

    # Status columns = everything that is not a metadata column.
    meta_names = {n for cands, _ in AGENT_STATUS_META_SPEC.values() for n in cands} | {"source_row"}
    candidate_status = [c for c in df.columns if c not in meta_names]
    df = resolve_columns(df, AGENT_STATUS_META_SPEC, ds, dq, expected_extra=set(candidate_status))
    if df["agent_id"].isna().all() and df["agent_name"].isna().all():
        raise DatasetError("Neither Agent Id nor Agent Name found - cannot identify agents.")

    status_cols, other_cols = [], []
    for c in candidate_status:
        vals = df[c].dropna()
        _, bad = parse_duration_seconds(vals, cfg.numeric_duration_unit)
        (status_cols if len(vals) == 0 or bad.mean() < 0.5 else other_cols).append(c)
    unmapped = [c for c in status_cols if c not in STATUS_DEFINITIONS]
    missing_expected = [STATUS_DEFINITIONS[k][0] for k in STATUS_DEFINITIONS
                        if k not in status_cols and k not in ("interacting", "communicating")]
    if unmapped:
        dq.add(ds, "Unexpected status columns (kept, group 'Unmapped (review)')", "WARNING",
               detail=", ".join(unmapped))
    if missing_expected:
        dq.add(ds, "Expected status columns not in file", "WARNING", detail=", ".join(missing_expected))
    if other_cols:
        dq.add(ds, "Non-duration extra columns kept as attributes", "INFO", detail=", ".join(other_cols))
    if not status_cols:
        raise DatasetError("No status duration columns found.")
    dq.add(ds, "Status columns detected", "INFO", detail=", ".join(status_cols))

    df = apply_skill_filter(df, cfg, ds, dq)
    flags: list[tuple[str, pd.Series]] = []
    df = add_interval_dimensions(df, cfg, ds, dq, flags)
    df["interval_complete"], bad = parse_bool(df["interval_complete"])
    dq.check(ds, "Unexpected Interval Complete values", bad, df)
    dq.check(ds, "Intervals not complete", ~df["interval_complete"].fillna(False).astype(bool), df,
             severity="INFO", example_cols=["report_date", "agent_name"])
    check_agent_identity(df, ds, dq)

    secs = {}
    for c in status_cols:
        s, bad = parse_duration_seconds(df[c], cfg.numeric_duration_unit)
        dq.check(ds, f"Invalid duration in status '{c}'", bad, df, severity="ERROR",
                 example_cols=["agent_name", c])
        flags.append((f"invalid_{c}", bad))
        if cfg.blank_status_duration_as_zero:
            s = s.where(df[c].notna() | bad, 0.0)
        secs[c] = s.round(3)
    sec = pd.DataFrame(secs, index=df.index)

    # --- Hierarchy checks (statuses overlap; see STATUS_DEFINITIONS) ---
    tol = cfg.duration_tolerance_seconds
    g = lambda c: sec[c] if c in sec else pd.Series(0.0, index=sec.index)  # noqa: E731
    if {"logged_in", "on_queue", "off_queue"} <= set(sec):
        m = (g("on_queue") + g("off_queue") - g("logged_in")).abs() > tol
        dq.check(ds, "On Queue + Off Queue differs from Logged In", m, df,
                 example_cols=["agent_name", "report_date"])
        flags.append(("on_plus_off_ne_logged_in", m))
    routing = [c for c in sec if c in STATUS_DEFINITIONS and STATUS_DEFINITIONS[c][1].startswith("3")]
    if "on_queue" in sec and routing:
        m = sec[routing].sum(axis=1) > g("on_queue") + tol
        dq.check(ds, f"{' + '.join(routing)} exceeds On Queue", m, df,
                 example_cols=["agent_name", "report_date"])
        flags.append(("routing_gt_on_queue", m))
    presence = [c for c in sec if c in STATUS_DEFINITIONS and STATUS_DEFINITIONS[c][1].startswith("4")]
    if "off_queue" in sec and presence:
        m = sec[presence].sum(axis=1) > g("off_queue") + tol
        dq.check(ds, "Sum of presence statuses exceeds Off Queue", m, df,
                 example_cols=["agent_name", "report_date"])
        flags.append(("presence_gt_off_queue", m))
        unexplained = (g("off_queue") - sec[presence].sum(axis=1)) > 60
        dq.check(ds, "Off Queue time not covered by the exported presence statuses (> 1 min)",
                 unexplained, df, severity="INFO",
                 detail="Likely presences not in the export (e.g. custom statuses). Not an error.",
                 example_cols=["agent_name"])
    if df["interval_minutes"].notna().any():
        too_long = sec.gt(df["interval_minutes"] * 60 + tol, axis=0).any(axis=1)
        dq.check(ds, "A status duration is longer than the interval", too_long, df, severity="ERROR",
                 example_cols=["agent_name", "report_date"])
        flags.append(("status_longer_than_interval", too_long))
    check_duplicates(df, ["interval_start", "interval_end", "agent_id"], ds, dq, flags)
    df["dq_flags"] = build_flags(df, flags)
    df = add_source_columns(df, path, load_ts)

    # --- One row per agent per day; each status becomes a column ---
    sec.columns = [f"{c}_seconds" for c in sec.columns]
    df = pd.concat([df.drop(columns=status_cols), sec], axis=1)
    daily = collapse_to_agent_day(df, list(sec.columns), ds, dq)
    daily = finalize_agent_status(daily)
    dq.check(ds, "Agents with the skill but no logged-in time (kept, agent_logged_in = FALSE)",
             ~daily["agent_logged_in"], daily, severity="INFO", example_cols=["agent_name"])
    dq.set_summary(ds, output_rows=len(daily), agents=daily["agent_id"].nunique(),
                   date_from=daily["report_date"].min(), date_to=daily["report_date"].max())
    return daily


def status_keys_in(df: pd.DataFrame) -> list[str]:
    """Status keys present as '<key>_seconds' columns, in display order."""
    keys = [c[: -len("_seconds")] for c in df.columns if c.endswith("_seconds")]
    known = [k for k in STATUS_DEFINITIONS if k in keys]
    return known + sorted(k for k in keys if k not in STATUS_DEFINITIONS)


def finalize_agent_status(df: pd.DataFrame) -> pd.DataFrame:
    """
    (Re)build the readable columns of the daily status sheet from the
    '<status>_seconds' columns: hh:mm:ss text, % of logged-in time and the
    agent_logged_in flag. Safe to run again after a history merge.
    """
    df = df.drop(columns=[c for c in ("interval_start_hour",) if c in df.columns])
    keys = status_keys_in(df)
    logged = df["logged_in_seconds"] if "logged_in_seconds" in df else df[[f"{k}_seconds" for k in keys]].max(axis=1)
    df["agent_logged_in"] = (logged.fillna(0) > 0).astype("boolean")
    status_cols = []
    for k in keys:
        sec_col = f"{k}_seconds"
        df[sec_col] = pd.to_numeric(df[sec_col], errors="coerce").round(3)
        df[f"{k}_hhmmss"] = seconds_to_hhmmss(df[sec_col])
        status_cols += [sec_col, f"{k}_hhmmss"]
        if k != "logged_in" and "logged_in_seconds" in df:
            df[f"{k}_pct_of_logged_in"] = (df[sec_col] / logged.where(logged > 0)).round(6)
            status_cols.append(f"{k}_pct_of_logged_in")
    first = ["report_date", "report_month", "weekday", "agent_id", "agent_name", "agent_logged_in",
             "interval_complete", "interval_start", "interval_end", "interval_minutes", "intervals_in_day"]
    first = [c for c in first if c in df.columns]
    tail_order = ("skills", "filters", "dq_flags", "source_file", "source_row", "load_timestamp")
    rest = [c for c in df.columns if c not in first and c not in status_cols and c not in tail_order]
    tail = [c for c in tail_order if c in df.columns]
    return df[first + status_cols + rest + tail]


def seconds_to_hhmmss(s: pd.Series) -> pd.Series:
    """12345.6 -> '03:25:46' (hours can exceed 24). Blank stays blank."""
    def f(v):
        if pd.isna(v):
            return None
        h, rem = divmod(int(round(float(v))), 3600)
        m, sec = divmod(rem, 60)
        return f"{h:02d}:{m:02d}:{sec:02d}"
    return s.map(f).astype(object)


def _join_unique(s: pd.Series) -> str | None:
    parts = dict.fromkeys(x.strip() for v in s.dropna().astype(str) for x in v.split(";") if x.strip())
    return "; ".join(parts) or None


def collapse_to_agent_day(df: pd.DataFrame, sum_cols: list[str], dataset: str,
                          dq: QualityReport) -> pd.DataFrame:
    """
    Return exactly one row per report_date + agent_id.

    The daily exports already have one row per agent, so this changes nothing
    today. If a file has several intervals per day (e.g. hourly), the numeric
    `sum_cols` are added up. Exact duplicate rows are counted once so time and
    calls are never double counted.
    """
    content = [c for c in df.columns if c not in ("source_row",)]
    exact = df.duplicated(subset=content, keep="first")
    if exact.any():
        dq.add(dataset, "Exact duplicate rows counted once in the daily totals", "WARNING", int(exact.sum()))
        df = df[~exact]
    key = ["report_date", "agent_id"]
    agg: dict = {c: (lambda x: x.sum(min_count=1)) for c in sum_cols}
    special = {
        "agent_name": "last", "interval_start": "min", "interval_end": "max",
        "interval_minutes": "sum", "load_timestamp": "max",
        "interval_complete": lambda x: bool(x.fillna(False).all()),
        "skills": _join_unique, "filters": _join_unique, "dq_flags": _join_unique,
        "source_file": _join_unique,
        "source_row": lambda x: ", ".join(str(int(v)) for v in x.dropna()),
    }
    for c in df.columns:
        if c in key or c in agg:
            continue
        agg[c] = special.get(c, "first")
    sizes = df.groupby(key, dropna=False).size()
    multi = int((sizes > 1).sum())
    if multi:
        dq.add(dataset, "Several intervals for the same agent and day combined into one daily row",
               "INFO", multi)
    else:
        dq.add(dataset, "One row per agent per day", "OK", 0)
    out = df.groupby(key, dropna=False, sort=False).agg(agg).reset_index()
    out["intervals_in_day"] = sizes.reindex(pd.MultiIndex.from_frame(out[key])).to_numpy()
    out["interval_complete"] = out["interval_complete"].astype("boolean")
    for c in sum_cols:
        if str(df[c].dtype) == "Int64":
            out[c] = out[c].astype("Int64")
    if "source_row" in out:
        out["source_row"] = out["source_row"].astype(object)
    return out


# =============================================================================
# 4. Agent performance
# =============================================================================
def process_agent_performance(path: Path, cfg: Config, dq: QualityReport, load_ts) -> pd.DataFrame:
    ds = "agent_performance"
    df = normalize_columns(read_csv_robust(path, cfg, dq, ds), ds)
    has_nr_count = any(c in df.columns for c in NOT_RESPONDING_COUNT_NAMES)
    df = resolve_columns(df, AGENT_PERFORMANCE_SPEC, ds, dq)
    if not has_nr_count:
        dq.add(ds, "No Not Responding count column (e.g. 'Alert - No Answer') in this export", "INFO",
               detail="not_responding_count is left blank (unknown), not 0.")
    if df["agent_id"].isna().all() and df["agent_name"].isna().all():
        raise DatasetError("Neither Agent Id nor Agent Name found - cannot identify agents.")

    df = apply_skill_filter(df, cfg, ds, dq)
    flags: list[tuple[str, pd.Series]] = []
    df = add_interval_dimensions(df, cfg, ds, dq, flags)
    df["interval_complete"], bad = parse_bool(df["interval_complete"])
    dq.check(ds, "Unexpected Interval Complete values", bad, df)
    check_agent_identity(df, ds, dq)

    count_cols = ["answered", "handled", "held", "transferred"] + (["not_responding_count"] if has_nr_count else [])
    for c in count_cols:
        df[c], bad = parse_count(df[c], cfg.blank_counts_as_zero)
        dq.check(ds, f"Invalid count in '{c}'", bad, df, example_cols=["agent_name"])
        flags.append((f"invalid_{c}", bad))

    df["avg_handle_seconds"], bad = parse_duration_seconds(df["avg_handle_raw"], cfg.numeric_duration_unit)
    df["avg_handle_seconds"] = df["avg_handle_seconds"].round(3)
    dq.check(ds, "Invalid Avg Handle duration", bad, df, severity="ERROR",
             example_cols=["agent_name", "avg_handle_raw"])
    flags.append(("invalid_avg_handle", bad))
    for c in ["avg_talk", "avg_hold", "avg_acw"]:
        s, bad = parse_duration_seconds(df[c], cfg.numeric_duration_unit)
        df[f"{c}_seconds"] = s.round(3)
        dq.check(ds, f"Invalid duration in '{c}'", bad, df, example_cols=["agent_name", c])
        flags.append((f"invalid_{c}", bad))
    df = df.drop(columns=["avg_talk", "avg_hold", "avg_acw"])

    # Avg Handle = total handle time / handled calls, so the total can be
    # rebuilt. Power BI can then compute a weighted AHT over any period:
    # SUM(total_handle_seconds) / SUM(handled).
    df["total_handle_seconds"] = (df["avg_handle_seconds"] * df["handled"].astype("float64")).round(3)
    df["has_activity"] = (df["answered"].fillna(0) > 0) | (df["handled"].fillna(0) > 0)

    m1 = (df["handled"] > 0) & df["avg_handle_seconds"].isna()
    m2 = df["avg_handle_seconds"].notna() & (df["handled"].fillna(0) == 0)
    m3 = (df["answered"] > 0) & (df["handled"].fillna(0) == 0)
    m4 = df["avg_handle_seconds"] > 4 * 3600
    dq.check(ds, "Handled > 0 but Avg Handle blank", m1, df, severity="ERROR", example_cols=["agent_name"])
    dq.check(ds, "Avg Handle present but Handled = 0", m2, df, example_cols=["agent_name", "avg_handle_raw"])
    dq.check(ds, "Answered > 0 but Handled = 0 (calls not yet completed?)", m3, df,
             example_cols=["agent_name", "answered"])
    dq.check(ds, "Avg Handle above 4 hours (check)", m4, df, example_cols=["agent_name", "avg_handle_raw"])
    ne = df["answered"] != df["handled"]
    dq.check(ds, "Answered differs from Handled", ne, df, severity="INFO",
             detail="Can be normal (transfers, calls answered in one interval and completed in another).",
             example_cols=["agent_name", "answered", "handled"])
    flags += [("handled_without_avg_handle", m1), ("avg_handle_without_handled", m2),
              ("answered_not_handled", m3), ("avg_handle_gt_4h", m4)]
    check_duplicates(df, ["interval_start", "interval_end", "agent_id"], ds, dq, flags)

    df["dq_flags"] = build_flags(df, flags)
    df = add_source_columns(df, path, load_ts)

    # --- One row per agent per day ---
    # Weighted helpers so averages stay correct if several intervals are
    # combined: talk and ACW weighted by handled calls, hold by held calls.
    h = df["handled"].astype("float64")
    df["_w_talk"] = df["avg_talk_seconds"] * h
    df["_w_acw"] = df["avg_acw_seconds"] * h
    df["_w_hold"] = df["avg_hold_seconds"] * df["held"].astype("float64")
    if not has_nr_count:
        df["not_responding_count"] = pd.array([pd.NA] * len(df), dtype="Int64")
    sums = ["answered", "handled", "held", "transferred", "not_responding_count", "total_handle_seconds",
            "_w_talk", "_w_acw", "_w_hold"]
    daily = collapse_to_agent_day(df, sums, ds, dq)
    multi = daily["intervals_in_day"] > 1
    hd = daily["handled"].astype("float64")
    hl = daily["held"].astype("float64")
    daily["avg_handle_seconds"] = (daily["total_handle_seconds"] / hd.where(hd > 0)).round(3)
    daily["avg_talk_seconds"] = daily["avg_talk_seconds"].where(~multi, daily["_w_talk"] / hd.where(hd > 0)).round(3)
    daily["avg_acw_seconds"] = daily["avg_acw_seconds"].where(~multi, daily["_w_acw"] / hd.where(hd > 0)).round(3)
    daily["avg_hold_seconds"] = daily["avg_hold_seconds"].where(~multi, daily["_w_hold"] / hl.where(hl > 0)).round(3)
    daily["avg_handle_raw"] = daily["avg_handle_raw"].where(~multi, None)
    daily = daily.drop(columns=["_w_talk", "_w_acw", "_w_hold"])
    daily = finalize_agent_performance(daily)
    dq.set_summary(ds, output_rows=len(daily), agents=daily["agent_id"].nunique(),
                   date_from=daily["report_date"].min(), date_to=daily["report_date"].max())
    return daily


def finalize_agent_performance(df: pd.DataFrame) -> pd.DataFrame:
    """Readable columns and column order for the daily performance sheet."""
    df["has_activity"] = ((df["answered"].fillna(0) > 0) | (df["handled"].fillna(0) > 0)).astype("boolean")
    for c in ("avg_handle", "avg_talk", "avg_hold", "avg_acw"):
        df[f"{c}_hhmmss"] = seconds_to_hhmmss(df[f"{c}_seconds"])
    return order_columns(df.drop(columns=[c for c in ("interval_start_hour",) if c in df]), [
        "report_date", "report_month", "weekday", "agent_id", "agent_name", "has_activity",
        "answered", "handled", "total_handle_seconds", "avg_handle_seconds", "avg_handle_hhmmss",
        "avg_handle_raw", "avg_talk_seconds", "avg_talk_hhmmss", "avg_hold_seconds",
        "avg_hold_hhmmss", "avg_acw_seconds", "avg_acw_hhmmss", "held", "transferred", "not_responding_count",
        "interval_complete", "interval_start", "interval_end", "interval_minutes",
        "intervals_in_day", "media_type", "skills", "filters"])


# =============================================================================
# Cross-dataset checks
# =============================================================================
def cross_checks(data: dict[str, pd.DataFrame], dq: QualityReport) -> None:
    ds = "cross_dataset"
    ranges = {k: (v["report_date"].min(), v["report_date"].max()) for k, v in data.items()
              if v is not None and len(v) and "report_date" in v}
    for k, (lo, hi) in ranges.items():
        dq.add(ds, f"Date range in output: {k}", "INFO", detail=f"{_fmt(lo)} to {_fmt(hi)}")
    if len(ranges) > 1:
        lo = max(r[0] for r in ranges.values())
        hi = min(r[1] for r in ranges.values())
        if lo > hi:
            dq.add(ds, "Datasets have no date in common", "WARNING")
        elif len({r for r in ranges.values()}) > 1:
            dq.add(ds, "Datasets cover different date ranges", "INFO",
                   detail=f"Common range: {_fmt(lo)} to {_fmt(hi)}. Compare KPIs across sheets only "
                          "for dates present in all of them.")

    sl, ab = data.get("service_level"), data.get("abandoned")
    perf, st = data.get("agent_performance"), data.get("agent_status")
    if sl is not None and len(sl):
        sl_day = sl.groupby("report_date")[["offered", "answered", "abandoned"]].sum()
        if ab is not None and len(ab):
            ab_day = ab.groupby("report_date").agg(records=("conversation_id", "count"),
                                                   abandoned_detail=("abandoned_flag", "sum"))
            j = sl_day.join(ab_day, how="inner")
            for d, r in j.iterrows():
                if r["records"] != r["offered"] or r["abandoned_detail"] != r["abandoned"]:
                    dq.add(ds, f"Queue totals vs call detail differ on {_fmt(d)}", "WARNING",
                           detail=f"Offered {r['offered']} vs {r['records']} conversations; "
                                  f"Abandoned {r['abandoned']} vs {r['abandoned_detail']}")
                else:
                    dq.add(ds, f"Queue totals match call detail on {_fmt(d)}", "OK",
                           detail=f"{r['records']} calls, {r['abandoned_detail']} abandoned")
            missing = sl_day[(sl_day["offered"] > 0) & ~sl_day.index.isin(ab_day.index)]
            if len(missing):
                dq.add(ds, "Days with offered calls but no rows in the call-detail file", "INFO",
                       len(missing), detail=", ".join(_fmt(d) for d in missing.index[:15]))
        if perf is not None and len(perf):
            p_day = perf.groupby("report_date")["answered"].sum()
            for d, v in p_day.items():
                if d in sl_day.index:
                    q = sl_day.loc[d, "answered"]
                    status = "OK" if v == q else "WARNING"
                    dq.add(ds, f"Answered: queue vs sum of agents on {_fmt(d)}", status,
                           detail=f"queue file {q}, agent file {v}" +
                           ("" if v == q else " - agent file may use another time zone/day boundary, "
                            "another filter, or agents outside the skill filter took calls."))
    if perf is not None and st is not None and len(perf) and len(st):
        logged = st.groupby(["report_date", "agent_id"])["logged_in_seconds"].sum()
        act = perf[perf["answered"] > 0].set_index(["report_date", "agent_id"])
        for (d, a), r in act.iterrows():
            if (d, a) in logged.index and logged.loc[(d, a)] == 0:
                dq.add(ds, "Agent answered calls but has no logged-in time", "WARNING",
                       detail=f"{r['agent_name']} on {_fmt(d)}")
        common = set(perf["report_date"].dropna()) & set(st["report_date"].dropna())
        for d in sorted(common):
            pa = set(perf.loc[perf["report_date"] == d, "agent_id"])
            sa = set(st.loc[st["report_date"] == d, "agent_id"])
            if pa != sa:
                dq.add(ds, f"Agent lists differ between status and performance on {_fmt(d)}", "INFO",
                       detail=f"{len(pa - sa)} only in performance, {len(sa - pa)} only in status")
            else:
                dq.add(ds, f"Same {len(pa)} agents in status and performance on {_fmt(d)}", "OK")
    if ab is not None and len(ab) and st is not None and len(st):
        names = set(st["agent_name"].dropna())
        users = {u.strip() for v in ab["users"].dropna() for u in v.split(";") if u.strip()}
        outside = sorted(users - names)
        if outside:
            dq.add(ds, "Agents in call detail 'Users' not in the filtered agent status data", "INFO",
                   len(outside), detail=", ".join(outside[:15]) +
                   " (other skills/transfers, or agents missing from the skill filter)")


# =============================================================================
# History merge and output
# =============================================================================
def _key_series(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    parts = []
    for c in cols:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            s = s.dt.strftime("%Y-%m-%d %H:%M:%S")
        parts.append(s.astype(object).where(s.notna(), "<NA>").astype(str).str.strip())
    key = parts[0]
    for p in parts[1:]:
        key = key + "|" + p
    return key


def _align_types(df: pd.DataFrame, template: pd.DataFrame) -> pd.DataFrame:
    """Cast columns read back from Excel to the dtypes of the freshly processed data."""
    for c in df.columns:
        if c not in template.columns:
            continue
        t = template[c].dtype
        try:
            if pd.api.types.is_datetime64_any_dtype(t):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            elif str(t) == "boolean":
                df[c] = df[c].map(lambda v: v if pd.isna(v) else str(v).strip().lower() in TRUE_VALUES
                                  ).astype("boolean")
            elif str(t) == "Int64":
                df[c] = pd.to_numeric(df[c], errors="coerce").round().astype("Int64")
            elif pd.api.types.is_float_dtype(t):
                df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64")
            elif pd.api.types.is_bool_dtype(t):
                df[c] = df[c].astype("boolean")
            else:
                df[c] = df[c].astype(object).where(df[c].notna(), None)
        except (ValueError, TypeError) as exc:
            log.warning("Could not align column %s to %s: %s", c, t, exc)
    return df


def merge_with_history(name: str, new: pd.DataFrame, old: pd.DataFrame | None,
                       dq: QualityReport) -> pd.DataFrame:
    """Upsert `new` into `old` by HISTORY_KEYS[name]; see module docstring."""
    if old is None or old.empty:
        dq.set_summary(name, history_rows_kept=0, history_rows_replaced=0)
        return new
    keys, comp = HISTORY_KEYS[name], COMPLETENESS_COLUMN[name]
    missing = [k for k in keys if k not in old.columns]
    if missing:
        raise DatasetError(f"Existing sheet {SHEETS[name]} lacks key columns {missing}; "
                           "run once with --no-history to rebuild it.")
    old = upgrade_legacy_sheet(name, old.copy())
    old = _align_types(old, new)
    new_cols, old_cols = set(new.columns), set(old.columns)
    if new_cols - old_cols:
        dq.add(name, "New columns not in history (blank for older rows)", "INFO",
               detail=", ".join(sorted(new_cols - old_cols)))
    if old_cols - new_cols:
        dq.add(name, "Historical columns not in new data (kept, blank for new rows)", "INFO",
               detail=", ".join(sorted(old_cols - new_cols)))

    k_old, k_new = _key_series(old, keys), _key_series(new, keys)
    overlap = set(k_old) & set(k_new)
    keep_old_for = set()
    if comp and comp in old.columns and comp in new.columns and overlap:
        old_complete = old.assign(_k=k_old).groupby("_k")[comp].agg(lambda s: bool(s.fillna(False).all()))
        new_complete = new.assign(_k=k_new).groupby("_k")[comp].agg(lambda s: bool(s.fillna(False).all()))
        keep_old_for = {k for k in overlap if old_complete.get(k, False) and not new_complete.get(k, True)}
    if keep_old_for:
        dq.add(name, "Incomplete new intervals NOT allowed to replace completed history", "INFO",
               int(k_new.isin(keep_old_for).sum()), detail=", ".join(sorted(keep_old_for)[:10]))
    replace = overlap - keep_old_for
    old_kept = old[~k_old.isin(replace)]
    new_kept = new[~k_new.isin(keep_old_for)]
    replaced = int(k_old.isin(replace).sum())
    dq.add(name, "History merge", "INFO",
           detail=f"{len(old)} historical rows: {len(old_kept)} kept, {replaced} replaced by newer export; "
                  f"{len(new_kept)} new rows added")
    dq.set_summary(name, history_rows_kept=len(old_kept), history_rows_replaced=replaced)
    frames = [f for f in (old_kept, new_kept) if len(f)]
    out = pd.concat(frames, ignore_index=True) if frames else new.iloc[0:0]
    return out[list(new.columns) + [c for c in out.columns if c not in new.columns]]


def upgrade_legacy_sheet(name: str, old: pd.DataFrame) -> pd.DataFrame:
    """Convert sheets written by the earlier version (agent x status rows) to the daily layout."""
    if name in ("agent_status", "agent_performance") and "interval_start_hour" in old.columns:
        old = old.drop(columns=["interval_start_hour"])
    if name != "agent_status" or not {"status", "duration_seconds"} <= set(old.columns):
        return old
    log.info("Converting existing agent_status_clean sheet from one-row-per-status to one-row-per-agent-day")
    to_key = {v[0]: k for k, v in STATUS_DEFINITIONS.items()}
    old["_k"] = old["status"].map(lambda v: to_key.get(v, normalize_name(v)))
    key = ["report_date", "agent_id"]
    pv = old.pivot_table(index=key, columns="_k", values="duration_seconds", aggfunc="sum")
    pv.columns = [f"{k}_seconds" for k in pv.columns]
    drop = {"status", "status_group", "status_sort", "duration_seconds", "duration_minutes",
            "duration_hours", "_k"}
    meta = old.drop(columns=[c for c in old.columns if c in drop]).drop_duplicates(key).set_index(key)
    return meta.join(pv).reset_index()


def load_existing_workbook(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        return {}
    try:
        return pd.read_excel(path, sheet_name=None, engine="openpyxl")
    except PermissionError as exc:
        raise SystemExit(f"Cannot read {path}: file is locked (close it in Excel). {exc}")
    except Exception as exc:  # noqa: BLE001 - any read failure must stop the run
        raise SystemExit(f"Existing workbook {path} could not be read ({exc}). Fix or move it, "
                         "or run with --no-history to rebuild from the supplied files only.")


SORT_KEYS = {
    "service_level": ["report_date", "queue_name", "interval_start"],
    "abandoned": ["call_datetime", "conversation_id"],
    "agent_status": ["report_date", "agent_name"],
    "agent_performance": ["report_date", "agent_name"],
}


def write_workbook(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    """
    Write the workbook in the system temp folder first, then move it into place,
    so a crash never leaves a half-written workbook. Nothing else is written to
    the output folder.
    """
    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", prefix="callcenter_")
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        _write_sheets(tmp, sheets)
        shutil.move(str(tmp), str(path))
    except PermissionError as exc:
        raise SystemExit(f"Cannot overwrite {path} - close it in Excel / Power BI and run again. {exc}")
    finally:
        if tmp.exists():
            tmp.unlink()


def _write_sheets(target: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(target, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm:ss",
                        date_format="yyyy-mm-dd") as xw:
        for sheet, df in sheets.items():
            out = df.copy()
            for c in out.columns:  # Excel cannot hold pandas NA objects / tz-aware values
                if str(out[c].dtype) in ("Int64", "boolean", "Float64"):
                    out[c] = out[c].astype(object).where(out[c].notna(), None)
            out.to_excel(xw, sheet_name=sheet, index=False)
            ws = xw.sheets[sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(out.columns, start=1):
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(max(len(str(col)) + 2, 12), 40)


# =============================================================================
# Orchestration
# =============================================================================
PROCESSORS = {
    "service_level": ("service_level_csv", "service_level_dir", process_service_level),
    "abandoned": ("abandoned_csv", "abandoned_dir", process_abandoned),
    "agent_status": ("agent_status_csv", "agent_status_dir", process_agent_status),
    "agent_performance": ("agent_performance_csv", "agent_performance_dir", process_agent_performance),
}


RUN_LOG_SHEET = "run_log"
DQ_LOG_SHEET = "data_quality_log"
RUN_LOG_COLUMNS = ["run_timestamp", "dataset", "status", "file_name", "file_path", "file_size_bytes",
                   "file_modified", "file_signature", "rows_in_file", "rows_in_sheet_after", "message"]
DQ_LOG_COLUMNS = ["run_timestamp", "dataset", "check", "status", "affected_rows", "detail", "examples"]


def file_signature(path: Path) -> str:
    """Size + modification time as text: a file replaced with new content is processed again."""
    st_ = path.stat()
    return f"{st_.st_size}|{st_.st_mtime_ns}"


def processed_files(run_log: pd.DataFrame | None) -> dict[str, dict[str, str]]:
    """{dataset: {absolute file path: signature}} of files already processed in earlier runs."""
    done: dict[str, dict[str, str]] = {}
    if run_log is None or run_log.empty or "status" not in run_log:
        return done
    ok = run_log[run_log["status"].astype(str) == "processed"]
    for row in ok.itertuples():
        done.setdefault(str(row.dataset), {})[str(row.file_path)] = str(row.file_signature)
    return done


def collect_inputs(name: str, cfg: Config, done: dict, dq: QualityReport, runlog: list) -> list[Path]:
    """Files to process for one export type: the single file (if given) + new files in the folder."""
    file_attr, dir_attr, _ = PROCESSORS[name]
    files: list[Path] = []
    single = getattr(cfg, file_attr)
    if single:
        files.append(Path(single))
    folder = getattr(cfg, dir_attr)
    if folder:
        d = Path(folder)
        if not d.is_dir():
            dq.add(name, "Input folder not found", "ERROR", detail=str(d))
            runlog.append({"dataset": name, "status": "folder not found", "file_path": str(d),
                           "message": "Check the folder path at the top of the script"})
            return files
        csvs = [f for f in d.iterdir()
                if f.is_file() and f.suffix.lower() == ".csv" and not f.name.startswith(("~$", "."))]
        csvs.sort(key=lambda f: (f.stat().st_mtime, f.name))   # oldest first, newest wins
        seen = done.get(name, {})
        new = [f for f in csvs if cfg.reprocess_all or seen.get(str(f.resolve())) != file_signature(f)]
        dq.add(name, f"Folder {d}", "INFO",
               detail=f"{len(csvs)} CSV files, {len(new)} new or changed, {len(csvs) - len(new)} already processed")
        files += new
        if not files:
            runlog.append({"dataset": name, "status": "no new files", "file_path": str(d),
                           "message": "Sheet kept unchanged"})
    return files


def run(cfg: Config) -> int:
    """Process all new input files and update the single output workbook. Returns an exit code."""
    cfg.output_dir.mkdir(parents=True, exist_ok=True)
    dq = QualityReport()
    run_ts = pd.Timestamp.now().floor("s")
    wb_path = cfg.output_dir / cfg.output_workbook_name
    if not cfg.append_history:
        cfg.reprocess_all = True          # rebuilding: every file in the folders is needed
    existing = load_existing_workbook(wb_path) if cfg.append_history else {}
    if existing:
        log.info("Existing workbook found: %s (%s)", wb_path,
                 ", ".join(f"{k}: {len(v)} rows" for k, v in existing.items()))
    old_runlog = existing.get(RUN_LOG_SHEET)
    old_dqlog = existing.get(DQ_LOG_SHEET)
    done = {} if cfg.reprocess_all else processed_files(old_runlog)

    configured = [k for k, (fa, da, _) in PROCESSORS.items() if getattr(cfg, fa) or getattr(cfg, da)]
    if not configured:
        log.error("No input folders or files set. Edit the folder paths at the top of callcenter_etl.py "
                  "or use --service-level-dir/--abandoned-dir/--agent-status-dir/--agent-performance-dir.")
        return 1

    final: dict[str, pd.DataFrame] = {}
    failed: list[str] = []
    runlog: list[dict] = []
    n_processed = 0
    for name, (_, _, func) in PROCESSORS.items():
        current = existing.get(SHEETS[name])
        for path in collect_inputs(name, cfg, done, dq, runlog):
            path = Path(path)
            entry = {"dataset": name, "file_name": path.name, "file_path": str(path.resolve())}
            if path.exists():
                st_ = path.stat()
                entry.update(file_size_bytes=st_.st_size, file_signature=file_signature(path),
                             file_modified=pd.Timestamp(st_.st_mtime, unit="s").floor("s"))
            log.info("=== Processing %s: %s ===", name, path)
            try:
                new = func(path, cfg, dq, run_ts)
                drop = [c for c in cfg.drop_columns.get(name, []) if c in new.columns]
                new = new.drop(columns=drop)
                merged = merge_with_history(name, new, current, dq)
                if name == "agent_status":
                    merged = finalize_agent_status(merged)
                elif name == "agent_performance":
                    merged = finalize_agent_performance(merged)
                current = merged.sort_values([c for c in SORT_KEYS[name] if c in merged.columns],
                                             kind="stable").reset_index(drop=True)
                entry.update(status="processed", rows_in_file=len(new), rows_in_sheet_after=len(current))
                n_processed += 1
            except DatasetError as exc:
                failed.append(f"{name}: {path.name}")
                dq.add(name, f"File could not be processed: {path.name}", "ERROR", detail=str(exc))
                entry.update(status="failed", message=str(exc)[:1000])
            runlog.append(entry)
        if current is not None:
            final[name] = current

    cross_checks({k: v for k, v in final.items()}, dq)

    # This run's logs are appended to the history kept inside the workbook.
    new_runlog = pd.DataFrame(runlog, columns=RUN_LOG_COLUMNS[1:])
    new_runlog.insert(0, "run_timestamp", run_ts)
    frames = [f for f in (old_runlog, new_runlog) if f is not None and len(f)]
    run_log = pd.concat(frames, ignore_index=True) if frames else new_runlog
    frames = [f for f in (old_dqlog, dq.log_rows(run_ts)) if f is not None and len(f)]
    dq_log = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=DQ_LOG_COLUMNS)

    sheets = {SHEETS[k]: final[k] for k in PROCESSORS if k in final}
    sheets[RUN_LOG_SHEET] = run_log[[c for c in RUN_LOG_COLUMNS if c in run_log.columns]]
    sheets[DQ_LOG_SHEET] = dq_log[[c for c in DQ_LOG_COLUMNS if c in dq_log.columns]]
    write_workbook(wb_path, sheets)
    log.info("Workbook written: %s", wb_path)
    n_err = sum(r["status"] == "ERROR" for r in dq.rows)
    n_warn = sum(r["status"] == "WARNING" for r in dq.rows)
    log.info("Finished: %d files processed, %d errors, %d warnings (details in the %s sheet)%s",
             n_processed, n_err, n_warn, DQ_LOG_SHEET, f"; FAILED: {', '.join(failed)}" if failed else "")
    folder_missing = any(r["check"] == "Input folder not found" for r in dq.rows)
    return 2 if failed or folder_missing else 0


def parse_args(argv: list[str] | None = None) -> Config:
    cfg = Config()
    p = argparse.ArgumentParser(description="Clean Genesys call-center exports for Power BI.")
    p.add_argument("--service-level-dir", help="Folder with queue performance CSVs (overrides SERVICE_LEVEL_FOLDER)")
    p.add_argument("--abandoned-dir", help="Folder with call detail CSVs (overrides ABANDONED_FOLDER)")
    p.add_argument("--agent-status-dir", help="Folder with agent status CSVs (overrides AGENT_STATUS_FOLDER)")
    p.add_argument("--agent-performance-dir", help="Folder with agent performance CSVs "
                                                   "(overrides AGENT_PERFORMANCE_FOLDER)")
    p.add_argument("--service-level", type=Path, help="A single queue performance CSV (file 1)")
    p.add_argument("--abandoned", type=Path, help="A single call detail / abandoned CSV (file 2)")
    p.add_argument("--agent-status", type=Path, help="A single agent status CSV (file 3)")
    p.add_argument("--agent-performance", type=Path, help="A single agent performance CSV (file 4)")
    p.add_argument("--files-only", action="store_true",
                   help="Ignore the folders set at the top of the script; process only the files given")
    p.add_argument("--reprocess-all", action="store_true",
                   help="Process every file in the folders again, even if processed before (no duplicates)")
    p.add_argument("--output-dir", type=Path, default=cfg.output_dir)
    p.add_argument("--output-name", default=cfg.output_workbook_name)
    p.add_argument("--skill", default=cfg.target_skill)
    p.add_argument("--skill-match", choices=["token", "exact"], default=cfg.skill_match_mode)
    p.add_argument("--no-history", action="store_true",
                   help="Ignore the existing workbook and rebuild it from the supplied files only")
    p.add_argument("--drop-caller-number", action="store_true",
                   help="Remove the caller phone number ('remote') from abandoned_calls_clean")
    a = p.parse_args(argv)
    if a.files_only:
        cfg.service_level_dir = cfg.abandoned_dir = cfg.agent_status_dir = cfg.agent_performance_dir = ""
    cfg.service_level_dir = a.service_level_dir if a.service_level_dir is not None else cfg.service_level_dir
    cfg.abandoned_dir = a.abandoned_dir if a.abandoned_dir is not None else cfg.abandoned_dir
    cfg.agent_status_dir = a.agent_status_dir if a.agent_status_dir is not None else cfg.agent_status_dir
    cfg.agent_performance_dir = (a.agent_performance_dir if a.agent_performance_dir is not None
                                 else cfg.agent_performance_dir)
    cfg.reprocess_all = a.reprocess_all
    cfg.service_level_csv = a.service_level or cfg.service_level_csv
    cfg.abandoned_csv = a.abandoned or cfg.abandoned_csv
    cfg.agent_status_csv = a.agent_status or cfg.agent_status_csv
    cfg.agent_performance_csv = a.agent_performance or cfg.agent_performance_csv
    cfg.output_dir = a.output_dir
    cfg.output_workbook_name = a.output_name
    cfg.target_skill = a.skill
    cfg.skill_match_mode = a.skill_match
    cfg.append_history = not a.no_history
    if a.drop_caller_number:
        cfg.drop_columns.setdefault("abandoned", []).append("remote")
    return cfg


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s", stream=sys.stdout)
    cfg = parse_args(argv)
    try:
        return run(cfg)
    except SystemExit as exc:
        log.error("%s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
