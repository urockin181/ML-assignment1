#!/usr/bin/env python3
"""
Generate the Power BI project (PBIP) "CallCenterDashboard" for the workbook
produced by ../callcenter_etl.py.

    python build_pbip.py            # writes ./CallCenterDashboard.pbip + folders

Output
------
CallCenterDashboard.pbip
CallCenterDashboard.SemanticModel/   TMDL model: 4 fact tables, Calendar, Agents, measures
CallCenterDashboard.Report/          PBIR report: "Queue Overview" and "Agents" pages

Open the .pbip in Power BI Desktop, then set the ExcelFilePath parameter
(Transform data > Edit parameters) to the full path of
callcenter_powerbi_clean.xlsx and click Refresh.

The report files use the PBIR format (report 1.3.0, page 1.4.0,
visualContainer 2.0.0), which Power BI Desktop has supported since mid-2025.
Re-run this script after changing columns or measures. It replaces the
generated folders.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import uuid
from pathlib import Path

HERE = Path(__file__).resolve().parent
NAME = "CallCenterDashboard"
DEFAULT_EXCEL_PATH = r"C:\CallCenter\output\callcenter_powerbi_clean.xlsx"

SL, AB, ST, AP = ("service_level_metrics_clean", "abandoned_calls_clean",
                  "agent_status_clean", "agent_performance_clean")

# Colors used for status categories
GREEN, RED, DARK_RED = "#107C10", "#C00000", "#8B0000"
CF_RED_BACK, CF_RED_STRONG = "#F4C7C3", "#E06666"
STATUS_COLORS = {
    "On Queue": "#1F6FB2", "Available": "#2E9E5B", "Break": "#8CC084",
    "Meeting": "#5BA3D9", "Training": "#8E7CC3",
    "Away": "#F28B82", "Meal": "#E06666", "Busy": "#C62828",
    "Other Off-Queue": "#8B0000", "Non-Allowed": "#C00000",
}


def uid(*parts: str) -> str:
    """Stable UUID so re-running the generator gives identical files."""
    return str(uuid.UUID(hashlib.md5("|".join(parts).encode()).hexdigest()))


def hid(*parts: str) -> str:
    """Stable 20-character id for pages and visuals."""
    return hashlib.md5("|".join(parts).encode()).hexdigest()[:20]


# =============================================================================
# Column schemas (must match the sheets written by callcenter_etl.py)
# =============================================================================
D, DT, I, F, B, S = "date", "datetime", "int", "double", "bool", "text"
TECH = {"filters", "source_file", "source_row", "load_timestamp"}  # hidden in the model

COMMON_TAIL = [("dq_flags", S), ("source_file", S), ("source_row", S), ("load_timestamp", DT)]

COLUMNS = {
    SL: [("report_date", D), ("report_month", S), ("weekday", S), ("interval_start", DT),
         ("interval_end", DT), ("interval_minutes", F), ("interval_start_hour", I),
         ("interval_complete", B), ("data_status", S), ("has_activity", B), ("queue_name", S),
         ("queue_id", S), ("media_type", S), ("offered", I), ("answered", I), ("abandoned", I),
         ("answered_within_sla", I), ("over_sla", I), ("handled", I), ("handled_in_sla", I),
         ("hold_count", I), ("transfer_count", I), ("service_level_pct_source", F),
         ("service_level_pct_calc", F), ("service_level_target_pct", F), ("answer_pct_source", F),
         ("answer_pct_calc", F), ("abandon_pct_source", F), ("abandon_pct_calc", F),
         ("asa_seconds", F), ("total_answer_wait_seconds", F), ("avg_wait_seconds", F),
         ("avg_handle_seconds", F), ("total_handle_seconds", F), ("avg_talk_seconds", F),
         ("avg_hold_seconds", F), ("avg_acw_seconds", F), ("filters", S)] + COMMON_TAIL,
    AB: [("conversation_id", S), ("call_datetime", DT), ("report_date", D), ("report_month", S),
         ("weekday", S), ("call_hour", I), ("half_hour_slot", S), ("is_abandoned", B),
         ("abandoned_flag", I), ("abandoned_raw", S), ("call_status", S),
         ("time_to_abandon_seconds", F), ("time_to_abandon_raw", S), ("duration_seconds", F),
         ("queue", S), ("direction", S), ("media_type", S), ("users", S), ("user_count", I),
         ("first_user", S), ("remote", S), ("wrap_up", S), ("full_export_completed", B),
         ("partial_result_timestamp", S), ("filters", S)] + COMMON_TAIL,
    AP: [("report_date", D), ("report_month", S), ("weekday", S), ("agent_id", S), ("agent_name", S),
         ("has_activity", B), ("answered", I), ("handled", I), ("total_handle_seconds", F),
         ("avg_handle_seconds", F), ("avg_handle_hhmmss", S), ("avg_handle_raw", S),
         ("avg_talk_seconds", F), ("avg_talk_hhmmss", S), ("avg_hold_seconds", F),
         ("avg_hold_hhmmss", S), ("avg_acw_seconds", F), ("avg_acw_hhmmss", S), ("held", I),
         ("transferred", I), ("interval_complete", B), ("interval_start", DT), ("interval_end", DT),
         ("interval_minutes", F), ("intervals_in_day", I), ("media_type", S), ("skills", S),
         ("filters", S)] + COMMON_TAIL,
}

# (key in the workbook, display name)
STATUSES = [("logged_in", "Logged In"), ("on_queue", "On Queue"), ("off_queue", "Off Queue"),
            ("idle", "Idle"), ("not_responding", "Not Responding"), ("available", "Available"),
            ("away", "Away"), ("break", "Break"), ("meal", "Meal"), ("busy", "Busy"),
            ("meeting", "Meeting"), ("training", "Training")]
_st_cols = [("report_date", D), ("report_month", S), ("weekday", S), ("agent_id", S),
            ("agent_name", S), ("agent_logged_in", B), ("interval_complete", B),
            ("interval_start", DT), ("interval_end", DT), ("interval_minutes", F),
            ("intervals_in_day", I)]
for k, _ in STATUSES:
    _st_cols += [(f"{k}_seconds", F), (f"{k}_hhmmss", S)]
    if k != "logged_in":
        _st_cols.append((f"{k}_pct_of_logged_in", F))
_st_cols += [("skills", S), ("filters", S)] + COMMON_TAIL
COLUMNS[ST] = _st_cols

# Agent status rules requested for the dashboard
ALLOWED = ["On Queue", "Idle", "Available", "Break", "Meeting", "Training"]
NON_ALLOWED = ["Not Responding", "Away", "Meal", "Busy", "Other Off-Queue"]


# =============================================================================
# Measures: (table, name, DAX, formatString or None, displayFolder, hidden)
# =============================================================================
def hhmmss(expr: str) -> str:
    """DAX that formats a number of seconds as hh:mm:ss text (hours may exceed 24)."""
    return (f"VAR _raw = {expr}\n"
            "VAR _s = ROUND ( _raw, 0 )\n"
            "RETURN\n"
            "    IF (\n"
            "        ISBLANK ( _raw ),\n"
            "        BLANK (),\n"
            "        FORMAT ( INT ( _s / 3600 ), \"00\" ) & \":\"\n"
            "            & FORMAT ( INT ( MOD ( _s, 3600 ) / 60 ), \"00\" ) & \":\"\n"
            "            & FORMAT ( MOD ( _s, 60 ), \"00\" )\n"
            "    )")


MEASURES: list[tuple] = []


def m(table, name, dax, fmt=None, folder=None, hidden=False):
    MEASURES.append((table, name, dax, fmt, folder, hidden))


# --- Queue (service level) ---
m(SL, "Offered Calls", f"SUM ( {SL}[offered] )", "#,0", "Volume")
m(SL, "Answered Calls", f"SUM ( {SL}[answered] )", "#,0", "Volume")
m(SL, "Abandoned Calls", f"SUM ( {SL}[abandoned] )", "#,0", "Volume")
m(SL, "Calls Within SLA", f"SUM ( {SL}[answered_within_sla] )", "#,0", "Service Level")
m(SL, "Service Level %", "DIVIDE ( [Calls Within SLA], [Answered Calls] )", "0.0%", "Service Level")
m(SL, "Abandon %", "DIVIDE ( [Abandoned Calls], [Offered Calls] )", "0.0%", "Volume")
m(SL, "Answer %", "DIVIDE ( [Answered Calls], [Offered Calls] )", "0.0%", "Volume")
m(SL, "ASA (sec)", f"DIVIDE ( SUM ( {SL}[total_answer_wait_seconds] ), [Answered Calls] )", "0", "Speed")
m(SL, "ASA", hhmmss("[ASA (sec)]"), None, "Speed")
m(SL, "Queue AHT (sec)", f"DIVIDE ( SUM ( {SL}[total_handle_seconds] ), SUM ( {SL}[handled] ) )", "0", "Speed")
m(SL, "Queue AHT", hhmmss("[Queue AHT (sec)]"), None, "Speed")
m(SL, "Service Level Target %", "0.80", "0%", "Targets")
m(SL, "Abandon Target %", "0.05", "0%", "Targets")
m(SL, "Service Level % Color",
  f"IF ( ISBLANK ( [Service Level %] ), BLANK (), IF ( [Service Level %] >= [Service Level Target %], \"{GREEN}\", \"{RED}\" ) )",
  None, "Formatting", True)
m(SL, "Abandon % Color",
  f"IF ( ISBLANK ( [Abandon %] ), BLANK (), IF ( [Abandon %] <= [Abandon Target %], \"{GREEN}\", \"{RED}\" ) )",
  None, "Formatting", True)

# --- Call detail (abandoned) ---
m(AB, "Call Records", f"COUNTROWS ( {AB} )", "#,0")
m(AB, "Abandoned Calls (Detail)", f"SUM ( {AB}[abandoned_flag] )", "#,0")
m(AB, "Abandon % (Detail)", "DIVIDE ( [Abandoned Calls (Detail)], [Call Records] )", "0.0%")
m(AB, "Avg Time to Abandon (sec)",
  f"CALCULATE ( AVERAGE ( {AB}[time_to_abandon_seconds] ), KEEPFILTERS ( {AB}[is_abandoned] = TRUE () ) )", "0")
m(AB, "Avg Time to Abandon", hhmmss("[Avg Time to Abandon (sec)]"))
m(AB, "Time to Abandon",
  hhmmss(f"CALCULATE ( SUM ( {AB}[time_to_abandon_seconds] ), KEEPFILTERS ( {AB}[is_abandoned] = TRUE () ) )"))

# --- Agent status ---
for key, label in STATUSES:
    m(ST, f"{label} Sec", f"SUM ( {ST}[{key}_seconds] )", "0", "Seconds", True)
presence = " + ".join(f"{ST}[{k}_seconds]" for k in
                      ("available", "away", "break", "meal", "busy", "meeting", "training"))
m(ST, "Other Off-Queue Sec",
  f"SUMX ( {ST}, MAX ( 0, {ST}[off_queue_seconds] - ( {presence} ) ) )", "0", "Seconds", True)
m(ST, "Allowed Sec",
  "[On Queue Sec] - [Not Responding Sec] + [Available Sec] + [Break Sec] + [Meeting Sec] + [Training Sec]",
  "0", "Seconds", True)
m(ST, "Non-Allowed Sec",
  "[Not Responding Sec] + [Away Sec] + [Meal Sec] + [Busy Sec] + [Other Off-Queue Sec]",
  "0", "Seconds", True)
for label in [s for _, s in STATUSES] + ["Other Off-Queue", "Allowed", "Non-Allowed"]:
    m(ST, f"{label} Hours", f"DIVIDE ( [{label} Sec], 3600 )", "0.00", "Hours")
    m(ST, f"{label} Time", hhmmss(f"[{label} Sec]"), None, "Time (hh:mm:ss)")
m(ST, "Allowed Status %", "DIVIDE ( [Allowed Sec], [Logged In Sec] )", "0.0%", "Compliance")
m(ST, "Non-Allowed Status %", "DIVIDE ( [Non-Allowed Sec], [Logged In Sec] )", "0.0%", "Compliance")
m(ST, "Allowed Status Target %", "0.90", "0%", "Targets")
m(ST, "Agents Logged In",
  f"CALCULATE ( DISTINCTCOUNT ( {ST}[agent_id] ), KEEPFILTERS ( {ST}[agent_logged_in] = TRUE () ) )", "#,0")
m(ST, "Agents With Skill", f"DISTINCTCOUNT ( {ST}[agent_id] )", "#,0")
for label in NON_ALLOWED:
    m(ST, f"{label} CF Color", f"IF ( [{label} Sec] > 0, \"{CF_RED_BACK}\" )", None, "Formatting", True)
m(ST, "Non-Allowed CF Color", f"IF ( [Non-Allowed Sec] > 0, \"{CF_RED_STRONG}\" )", None, "Formatting", True)
m(ST, "Allowed Status % Color",
  f"IF ( ISBLANK ( [Allowed Status %] ), BLANK (), IF ( [Allowed Status %] >= [Allowed Status Target %], \"{GREEN}\", \"{RED}\" ) )",
  None, "Formatting", True)

# --- Agent performance ---
m(AP, "Answered (Agents)", f"SUM ( {AP}[answered] )", "#,0")
m(AP, "Handled (Agents)", f"SUM ( {AP}[handled] )", "#,0")
m(AP, "Held Calls", f"SUM ( {AP}[held] )", "#,0")
m(AP, "Transferred Calls", f"SUM ( {AP}[transferred] )", "#,0")
m(AP, "AHT (sec)", f"DIVIDE ( SUM ( {AP}[total_handle_seconds] ), [Handled (Agents)] )", "0", "Handle Time")
m(AP, "AHT", hhmmss("[AHT (sec)]"), None, "Handle Time")
m(AP, "Avg Talk (sec)", f"DIVIDE ( SUMX ( {AP}, {AP}[avg_talk_seconds] * {AP}[handled] ), [Handled (Agents)] )",
  "0", "Handle Time")
m(AP, "Avg Talk", hhmmss("[Avg Talk (sec)]"), None, "Handle Time")
m(AP, "Avg Hold (sec)", f"DIVIDE ( SUMX ( {AP}, {AP}[avg_hold_seconds] * {AP}[held] ), [Held Calls] )",
  "0", "Handle Time")
m(AP, "Avg Hold", hhmmss("[Avg Hold (sec)]"), None, "Handle Time")
m(AP, "Avg ACW (sec)", f"DIVIDE ( SUMX ( {AP}, {AP}[avg_acw_seconds] * {AP}[handled] ), [Handled (Agents)] )",
  "0", "Handle Time")
m(AP, "Avg ACW", hhmmss("[Avg ACW (sec)]"), None, "Handle Time")


# =============================================================================
# Semantic model (TMDL)
# =============================================================================
TMDL_TYPES = {D: "dateTime", DT: "dateTime", I: "int64", F: "double", B: "boolean", S: "string"}
M_TYPES = {D: "type date", DT: "type datetime", I: "Int64.Type", F: "type number",
           B: "type logical", S: "type text"}
FORMATS = {D: "yyyy-mm-dd", DT: "General Date", I: "0"}


def tmdl_name(n: str) -> str:
    return n if n.replace("_", "").isalnum() and not n[0].isdigit() else "'" + n.replace("'", "''") + "'"


def indent(text: str, tabs: int) -> str:
    return "\n".join(("\t" * tabs + ln) if ln.strip() else "" for ln in text.splitlines())


def excel_source(sheet: str, cols: list[tuple[str, str]]) -> str:
    names = ", ".join(f'"{c}"' for c, _ in cols)
    types = ",\n".join(f'        {{"{c}", {M_TYPES[t]}}}' for c, t in cols)
    return (
        "let\n"
        "    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),\n"
        f'    Sheet = Source{{[Item = "{sheet}", Kind = "Sheet"]}}[Data],\n'
        "    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),\n"
        f"    Selected = Table.SelectColumns(Promoted, {{{names}}}, MissingField.UseNull),\n"
        "    Typed = Table.TransformColumnTypes(\n"
        "        Selected,\n"
        "        {\n"
        f"{types}\n"
        "        }\n"
        "    )\n"
        "in\n"
        "    Typed")


def column_block(table: str, name: str, typ: str, hidden: bool = False, extra: list[str] | None = None) -> str:
    lines = [f"\tcolumn {tmdl_name(name)}", f"\t\tdataType: {TMDL_TYPES[typ]}"]
    if hidden:
        lines.append("\t\tisHidden")
    if typ in FORMATS:
        lines.append(f"\t\tformatString: {FORMATS[typ]}")
    elif typ == F:
        lines.append("\t\tformatString: 0.00")
    lines += [f"\t\tlineageTag: {uid(table, 'col', name)}", "\t\tsummarizeBy: none",
              f"\t\tsourceColumn: {name}"]
    lines += [f"\t\t{e}" for e in (extra or [])]
    lines.append("")
    if typ == D:
        lines += ["\t\tannotation UnderlyingDateTimeDataType = Date", ""]
    lines += ["\t\tannotation SummarizationSetBy = User", ""]
    return "\n".join(lines)


def measure_block(table, name, dax, fmt, folder, hidden) -> str:
    lines = [f"\tmeasure {tmdl_name(name)} = ```", indent(dax, 3), "\t\t\t```"]
    if fmt:
        lines.append(f"\t\tformatString: {fmt}")
    if hidden:
        lines.append("\t\tisHidden")
    if folder:
        lines.append(f"\t\tdisplayFolder: {folder}")
    lines += [f"\t\tlineageTag: {uid(table, 'measure', name)}", ""]
    return "\n".join(lines)


def table_tmdl(table: str, cols, source_m: str, extra_table_props: list[str] | None = None,
               column_extras: dict | None = None, hidden_cols: set | None = None) -> str:
    out = [f"table {tmdl_name(table)}", f"\tlineageTag: {uid(table)}"]
    out += [f"\t{p}" for p in (extra_table_props or [])]
    out.append("")
    for mt in MEASURES:
        if mt[0] == table:
            out.append(measure_block(*mt))
    for c, t in cols:
        hidden = c in TECH or c in (hidden_cols or set())
        out.append(column_block(table, c, t, hidden, (column_extras or {}).get(c)))
    out += [f"\tpartition {tmdl_name(table)} = m", "\t\tmode: import", "\t\tsource =",
            indent(source_m, 4), "", "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(out)


CALENDAR_M = """let
    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),
    SheetDates = (sheet as text) as list =>
        let
            Found = Table.SelectRows(Source, each [Item] = sheet and [Kind] = "Sheet"),
            Data = if Table.IsEmpty(Found) then #table({"report_date"}, {})
                   else Table.PromoteHeaders(Found{0}[Data], [PromoteAllScalars = true]),
            Values = if Table.HasColumns(Data, "report_date") then Table.Column(Data, "report_date") else {}
        in
            List.RemoveNulls(List.Transform(Values, each try Date.From(_) otherwise null)),
    AllDates = List.Combine({
        SheetDates("service_level_metrics_clean"), SheetDates("abandoned_calls_clean"),
        SheetDates("agent_status_clean"), SheetDates("agent_performance_clean")}),
    Today = Date.From(DateTime.LocalNow()),
    StartDate = Date.StartOfMonth(if List.IsEmpty(AllDates) then Today else List.Min(AllDates)),
    EndDate = Date.EndOfMonth(if List.IsEmpty(AllDates) then Today else List.Max(AllDates)),
    Dates = List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)),
    Base = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),
    Typed = Table.TransformColumnTypes(Base, {{"Date", type date}}),
    AddYear = Table.AddColumn(Typed, "Year", each Date.Year([Date]), Int64.Type),
    AddMonthNo = Table.AddColumn(AddYear, "Month Sort", each Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type),
    AddMonth = Table.AddColumn(AddMonthNo, "Month", each Date.ToText([Date], "MMM yyyy", "en-US"), type text),
    AddDay = Table.AddColumn(AddMonth, "Day", each Date.Day([Date]), Int64.Type),
    AddWeekdayNo = Table.AddColumn(AddDay, "Weekday No", each Date.DayOfWeek([Date], Day.Monday) + 1, Int64.Type),
    AddWeekday = Table.AddColumn(AddWeekdayNo, "Weekday", each Date.ToText([Date], "ddd", "en-US"), type text)
in
    AddWeekday"""

AGENTS_M = """let
    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),
    SheetAgents = (sheet as text) as table =>
        let
            Found = Table.SelectRows(Source, each [Item] = sheet and [Kind] = "Sheet"),
            Data = if Table.IsEmpty(Found) then #table({"agent_id", "agent_name", "report_date"}, {})
                   else Table.PromoteHeaders(Found{0}[Data], [PromoteAllScalars = true])
        in
            Table.SelectColumns(Data, {"agent_id", "agent_name", "report_date"}, MissingField.UseNull),
    Combined = Table.Combine({SheetAgents("agent_status_clean"), SheetAgents("agent_performance_clean")}),
    Typed = Table.TransformColumnTypes(Combined, {{"agent_id", type text}, {"agent_name", type text}, {"report_date", type date}}),
    NotBlank = Table.SelectRows(Typed, each [agent_id] <> null and [agent_id] <> ""),
    Sorted = Table.Buffer(Table.Sort(NotBlank, {{"report_date", Order.Descending}})),
    OnePerAgent = Table.Distinct(Sorted, {"agent_id"}),
    Result = Table.SelectColumns(OnePerAgent, {"agent_id", "agent_name"})
in
    Result"""


def build_model(root: Path) -> None:
    sm = root / f"{NAME}.SemanticModel"
    d = sm / "definition"
    (d / "tables").mkdir(parents=True)
    write_json(sm / "definition.pbism", {"version": "4.0", "settings": {}})
    write_json(sm / ".platform", platform("SemanticModel"))
    (d / "database.tmdl").write_text("database\n\tcompatibilityLevel: 1601\n\n", encoding="utf-8")

    tables = [SL, AB, ST, AP, "Calendar", "Agents"]
    query_order = json.dumps(["ExcelFilePath"] + tables)
    (d / "model.tmdl").write_text(
        "model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
        "\tsourceQueryCulture: en-US\n\tdataAccessOptions\n\t\tlegacyRedirects\n\t\treturnErrorValuesAsNull\n\n"
        f"annotation __PBI_TimeIntelligenceEnabled = 0\n\nannotation PBI_QueryOrder = {query_order}\n\n"
        + "".join(f"ref table {tmdl_name(t)}\n" for t in tables), encoding="utf-8")

    (d / "expressions.tmdl").write_text(
        "/// Full path to callcenter_powerbi_clean.xlsx (output of callcenter_etl.py)\n"
        "expression ExcelFilePath =\n"
        f'\t\t"{DEFAULT_EXCEL_PATH}" meta [\n'
        '\t\t\tIsParameterQuery = true,\n\t\t\tType = "Text",\n\t\t\tIsParameterQueryRequired = true\n\t\t]\n'
        f"\tlineageTag: {uid('ExcelFilePath')}\n\n\tannotation PBI_ResultType = Text\n", encoding="utf-8")

    for t in (SL, AB, ST, AP):
        hidden = {"agent_id"} if t in (ST, AP) else set()
        (d / "tables" / f"{t}.tmdl").write_text(
            table_tmdl(t, COLUMNS[t], excel_source(t, COLUMNS[t]), hidden_cols=hidden), encoding="utf-8")

    cal_cols = [("Date", D), ("Year", I), ("Month Sort", I), ("Month", S), ("Day", I),
                ("Weekday No", I), ("Weekday", S)]
    (d / "tables" / "Calendar.tmdl").write_text(table_tmdl(
        "Calendar", cal_cols, CALENDAR_M, extra_table_props=["dataCategory: Time"],
        column_extras={"Date": ["isKey"], "Month": ["sortByColumn: 'Month Sort'"],
                       "Weekday": ["sortByColumn: 'Weekday No'"]},
        hidden_cols={"Month Sort", "Weekday No"}), encoding="utf-8")
    (d / "tables" / "Agents.tmdl").write_text(table_tmdl(
        "Agents", [("agent_id", S), ("agent_name", S)], AGENTS_M,
        column_extras={"agent_id": ["isKey"]}), encoding="utf-8")

    rels = []
    for t in (SL, AB, ST, AP):
        rels.append((f"{t}.report_date", "Calendar.Date"))
    for t in (ST, AP):
        rels.append((f"{t}.agent_id", "Agents.agent_id"))
    (d / "relationships.tmdl").write_text("".join(
        f"relationship {uid('rel', f, to)}\n\tfromColumn: {f}\n\ttoColumn: {to}\n\n" for f, to in rels),
        encoding="utf-8")


# =============================================================================
# Report (PBIR)
# =============================================================================
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
VC_SCHEMA = f"{SCHEMA}/visualContainer/2.0.0/schema.json"


def lit(v: str) -> dict:
    return {"expr": {"Literal": {"Value": v}}}


def q(s: str) -> str:
    return "'" + s.replace("'", "''") + "'"


def color(hexv: str) -> dict:
    return {"solid": {"color": lit(q(hexv))}}


def col(entity: str, prop: str) -> dict:
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def mea(entity: str, prop: str) -> dict:
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def ref(field: dict) -> str:
    k = "Column" if "Column" in field else "Measure"
    return f"{field[k]['Expression']['SourceRef']['Entity']}.{field[k]['Property']}"


def proj(field: dict, display: str | None = None, category: bool = False) -> dict:
    p = {"field": field, "queryRef": ref(field)}
    if display:
        p["displayName"] = display
    if category:
        p["active"] = True
    return p


def measure_color(field: dict) -> dict:
    return {"solid": {"color": {"expr": field}}}


class Page:
    def __init__(self, name: str, display: str, height: int):
        self.id, self.display, self.height, self.visuals = hid("page", name), display, height, []

    def add(self, key: str, x, y, w, h, visual: dict, title: str | None = None, background=True):
        vco = {}
        if title:
            vco["title"] = [{"properties": {"show": lit("true"), "text": lit(q(title))}}]
        if not background:
            vco["background"] = [{"properties": {"show": lit("false")}}]
        if vco:
            visual["visualContainerObjects"] = vco
        visual.setdefault("drillFilterOtherVisuals", True)
        n = len(self.visuals)
        self.visuals.append({
            "$schema": VC_SCHEMA, "name": hid(self.id, key),
            "position": {"x": x, "y": y, "z": n * 1000, "height": h, "width": w, "tabOrder": n * 1000},
            "visual": visual})


def textbox(text: str, size="18pt", color_hex="#252423") -> dict:
    return {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": [
        {"textRuns": [{"value": text, "textStyle": {"fontSize": size, "fontWeight": "bold",
                                                   "color": color_hex}}]}]}}]}}


def slicer(field: dict, mode: str) -> dict:
    return {"visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [proj(field, category=True)]}}},
            "objects": {"data": [{"properties": {"mode": lit(q(mode))}}]}}


def card(field: dict, display: str | None = None, value_color: str | None = None) -> dict:
    v = {"visualType": "card",
         "query": {"queryState": {"Values": {"projections": [proj(field, display)]}}},
         "objects": {"labels": [{"properties": {"fontSize": lit("20D")}}]}}
    if value_color:
        v["objects"]["labels"][0]["properties"]["color"] = color(value_color)
    return v


def series_colors(fields_colors: list[tuple[dict, str]]) -> list:
    return [{"properties": {"fill": color(c)}, "selector": {"metadata": ref(f)}} for f, c in fields_colors]


def chart(vtype: str, category: dict, y: list[tuple[dict, str | None]], y2=None, sort=None,
          legend=True, display=None) -> dict:
    qs = {"Category": {"projections": [proj(category, display, category=True)]},
          "Y": {"projections": [proj(f) for f, _ in y]}}
    if y2:
        qs["Y2"] = {"projections": [proj(f) for f, _ in y2]}
    v = {"visualType": vtype, "query": {"queryState": qs}, "objects": {}}
    if sort:
        v["query"]["sortDefinition"] = {"sort": [{"field": sort[0], "direction": sort[1]}]}
    colors = [(f, c) for f, c in y + (y2 or []) if c]
    if colors:
        v["objects"]["dataPoint"] = series_colors(colors)
    if legend:
        v["objects"]["legend"] = [{"properties": {"show": lit("true"), "position": lit("'Top'")}}]
    return v


def table(cols: list[tuple[dict, str | None]], sort=None, cf: list[tuple[dict, str, dict]] | None = None) -> dict:
    """cf entries: (formatted field, property 'backColor'|'fontColor', color measure field)."""
    v = {"visualType": "tableEx",
         "query": {"queryState": {"Values": {"projections": [proj(f, d) for f, d in cols]}}},
         "objects": {"values": []}}
    if sort:
        v["query"]["sortDefinition"] = {"sort": [{"field": sort[0], "direction": sort[1]}]}
    for target, prop, color_measure in cf or []:
        v["objects"]["values"].append({
            "properties": {prop: measure_color(color_measure)},
            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": ref(target)}})
    if not v["objects"]["values"]:
        del v["objects"]
    return v


def build_pages() -> list[Page]:
    W = 1280
    # ---------------- Page 1: Queue overview ----------------
    p1 = Page("queue", "Queue Overview", 900)
    p1.add("title", 16, 8, 700, 50, textbox("Queue Overview - GL AUDI EHD"), background=False)
    p1.add("date", 900, 6, 364, 60, slicer(col("Calendar", "Date"), "Between"), title="Date")
    kpis = [(mea(SL, "Offered Calls"), None, None), (mea(SL, "Answered Calls"), None, None),
            (mea(SL, "Abandoned Calls"), None, RED), (mea(SL, "Service Level %"), None, None),
            (mea(SL, "Abandon %"), None, None), (mea(SL, "Calls Within SLA"), None, None),
            (mea(SL, "ASA"), "Avg Speed of Answer", None),
            (mea(AB, "Avg Time to Abandon"), "Avg Time to Abandon", None)]
    cw = (W - 32 - 7 * 8) / 8
    for i, (f, disp, c) in enumerate(kpis):
        p1.add(f"kpi{i}", 16 + i * (cw + 8), 74, cw, 92, card(f, disp, c))
    p1.add("trend", 16, 176, 820, 310, chart(
        "lineClusteredColumnComboChart", col("Calendar", "Date"),
        [(mea(SL, "Offered Calls"), "#A6A6A6"), (mea(SL, "Answered Calls"), "#1F6FB2"),
         (mea(SL, "Abandoned Calls"), RED)],
        y2=[(mea(SL, "Service Level %"), GREEN), (mea(SL, "Abandon %"), "#E8A33D")]),
        title="Daily volume, Service Level % and Abandon %")
    p1.add("byhour", 846, 176, 418, 310, chart(
        "clusteredColumnChart", col(AB, "call_hour"),
        [(mea(AB, "Call Records"), "#1F6FB2"), (mea(AB, "Abandoned Calls (Detail)"), RED)],
        sort=(col(AB, "call_hour"), "Ascending"), display="Hour of day"),
        title="Calls and abandoned calls by hour (call detail)")
    sl_pct, ab_pct = mea(SL, "Service Level %"), mea(SL, "Abandon %")
    p1.add("daily", 16, 496, 760, 390, table(
        [(col("Calendar", "Date"), "Date"), (mea(SL, "Offered Calls"), "Offered"),
         (mea(SL, "Answered Calls"), "Answered"), (mea(SL, "Abandoned Calls"), "Abandoned"),
         (mea(SL, "Calls Within SLA"), "Within SLA"), (sl_pct, "Service Level %"),
         (ab_pct, "Abandon %"), (mea(SL, "ASA"), "ASA"), (mea(SL, "Queue AHT"), "AHT"),
         (mea(AB, "Call Records"), "Detail: calls"), (mea(AB, "Abandoned Calls (Detail)"), "Detail: abandoned")],
        sort=(col("Calendar", "Date"), "Ascending"),
        cf=[(sl_pct, "fontColor", mea(SL, "Service Level % Color")),
            (ab_pct, "fontColor", mea(SL, "Abandon % Color"))]),
        title="Daily summary (red = below Service Level target / above Abandon target)")
    p1.add("abandoned", 786, 496, 478, 390, table(
        [(col(AB, "call_datetime"), "Call time"), (mea(AB, "Time to Abandon"), "Waited before abandoning"),
         (col(AB, "users"), "Agent(s) alerted"), (col(AB, "conversation_id"), "Conversation ID")],
        sort=(col(AB, "call_datetime"), "Ascending")),
        title="Abandoned calls")

    # ---------------- Page 2: Agents ----------------
    p2 = Page("agents", "Agents", 1110)
    p2.add("title", 16, 8, 640, 50, textbox("Agents - GL AUDI EHD"), background=False)
    p2.add("date", 670, 6, 300, 60, slicer(col("Calendar", "Date"), "Between"), title="Date")
    p2.add("agent", 980, 6, 284, 60, slicer(col("Agents", "agent_name"), "Dropdown"), title="Agent")
    kpis2 = [(mea(ST, "Agents Logged In"), None, None), (mea(ST, "Logged In Time"), "Logged In (hh:mm:ss)", None),
             (mea(ST, "Allowed Status %"), "Allowed Status %", GREEN),
             (mea(ST, "Non-Allowed Time"), "Non-Allowed Status Time", RED),
             (mea(AP, "Answered (Agents)"), "Answered", None), (mea(AP, "Handled (Agents)"), "Handled", None),
             (mea(AP, "AHT"), "Avg Handle Time", None)]
    cw = (W - 32 - 6 * 8) / 7
    for i, (f, disp, c) in enumerate(kpis2):
        p2.add(f"kpi{i}", 16 + i * (cw + 8), 74, cw, 92, card(f, disp, c))
    stack = [("On Queue", "On Queue"), ("Available", "Available"), ("Break", "Break"),
             ("Meeting", "Meeting"), ("Training", "Training"), ("Away", "Away"), ("Meal", "Meal"),
             ("Busy", "Busy"), ("Other Off-Queue", "Other Off-Queue")]
    p2.add("stack", 16, 176, 830, 340, chart(
        "barChart", col("Agents", "agent_name"),
        [(mea(ST, f"{s} Hours"), STATUS_COLORS[c]) for s, c in stack],
        sort=(col("Agents", "agent_name"), "Ascending"), display="Agent"),
        title="Status hours by agent (red shades = statuses outside the allowed list)")
    donut = {"visualType": "donutChart", "query": {"queryState": {"Y": {"projections": [
        proj(mea(ST, f"{s} Hours")) for s in ("On Queue", "Available", "Break", "Meeting", "Training", "Non-Allowed")]}}},
        "objects": {"dataPoint": series_colors([(mea(ST, f"{s} Hours"), STATUS_COLORS[s]) for s in
                                                ("On Queue", "Available", "Break", "Meeting", "Training", "Non-Allowed")]),
                    "legend": [{"properties": {"show": lit("true"), "position": lit("'Bottom'")}}]}}
    p2.add("donut", 856, 176, 408, 340, donut, title="Status distribution (hours)")

    st_cols = [(col("Agents", "agent_name"), "Agent"), (mea(ST, "Logged In Time"), "Logged In")]
    st_cols += [(mea(ST, f"{s} Time"), s) for s in ("On Queue", "Idle", "Available", "Break", "Meeting", "Training")]
    st_cols += [(mea(ST, f"{s} Time"), s) for s in NON_ALLOWED]
    st_cols += [(mea(ST, "Non-Allowed Time"), "Non-Allowed total"), (mea(ST, "Allowed Status %"), "Allowed %")]
    cf = [(mea(ST, f"{s} Time"), "backColor", mea(ST, f"{s} CF Color")) for s in NON_ALLOWED]
    cf += [(mea(ST, "Non-Allowed Time"), "backColor", mea(ST, "Non-Allowed CF Color")),
           (mea(ST, "Allowed Status %"), "fontColor", mea(ST, "Allowed Status % Color"))]
    p2.add("status", 16, 526, 1248, 330, table(st_cols, sort=(col("Agents", "agent_name"), "Ascending"), cf=cf),
           title="Agent status time (hh:mm:ss). Red = time in a status outside "
                 "On Queue / Idle / Available / Break / Meeting / Training")
    p2.add("perf", 16, 866, 820, 230, table(
        [(col("Agents", "agent_name"), "Agent"), (mea(AP, "Answered (Agents)"), "Answered"),
         (mea(AP, "Handled (Agents)"), "Handled"), (mea(AP, "AHT"), "AHT"),
         (mea(AP, "Avg Talk"), "Avg Talk"), (mea(AP, "Avg Hold"), "Avg Hold"),
         (mea(AP, "Avg ACW"), "Avg ACW"), (mea(AP, "Held Calls"), "Held"),
         (mea(AP, "Transferred Calls"), "Transferred")],
        sort=(col("Agents", "agent_name"), "Ascending")),
        title="Agent performance (AHT = total handle time / handled calls)")
    p2.add("answered", 846, 866, 418, 230, chart(
        "clusteredBarChart", col("Agents", "agent_name"),
        [(mea(AP, "Answered (Agents)"), "#1F6FB2"), (mea(AP, "Handled (Agents)"), "#5BA3D9")],
        sort=(mea(AP, "Answered (Agents)"), "Descending"), display="Agent"),
        title="Answered and handled calls by agent")
    return [p1, p2]


def platform(kind: str) -> dict:
    return {"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": kind, "displayName": NAME},
            "config": {"version": "2.0", "logicalId": uid(NAME, kind)}}


def build_report(root: Path, theme_src: Path) -> None:
    rp = root / f"{NAME}.Report"
    d = rp / "definition"
    write_json(rp / "definition.pbir", {"version": "4.0", "datasetReference": {
        "byPath": {"path": f"../{NAME}.SemanticModel"}}})
    write_json(rp / ".platform", platform("Report"))
    write_json(d / "version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    theme_dir = rp / "StaticResources" / "SharedResources" / "BaseThemes"
    theme_dir.mkdir(parents=True)
    shutil.copy(theme_src, theme_dir / "CY24SU10.json")
    write_json(d / "report.json", {
        "$schema": f"{SCHEMA}/report/1.3.0/schema.json",
        "themeCollection": {"baseTheme": {"name": "CY24SU10", "reportVersionAtImport": "5.61",
                                          "type": "SharedResources"}},
        "layoutOptimization": "None",
        "resourcePackages": [{"name": "SharedResources", "type": "SharedResources", "items": [
            {"name": "CY24SU10", "path": "BaseThemes/CY24SU10.json", "type": "BaseTheme"}]}],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarizedAndUnderlying",
                     "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True,
                     "useEnhancedTooltips": True}})
    pages = build_pages()
    write_json(d / "pages" / "pages.json", {"$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
                                            "pageOrder": [p.id for p in pages], "activePageName": pages[0].id})
    for p in pages:
        pd_ = d / "pages" / p.id
        write_json(pd_ / "page.json", {"$schema": f"{SCHEMA}/page/1.4.0/schema.json", "name": p.id,
                                       "displayName": p.display, "displayOption": "FitToWidth",
                                       "height": p.height, "width": 1280})
        for v in p.visuals:
            write_json(pd_ / "visuals" / v["name"] / "visual.json", v)


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def main() -> None:
    theme = HERE / "theme" / "CY24SU10.json"
    if not theme.is_file():
        raise SystemExit(f"Base theme file missing: {theme}")
    for p in (HERE / f"{NAME}.SemanticModel", HERE / f"{NAME}.Report"):
        if p.exists():
            shutil.rmtree(p)
    write_json(HERE / f"{NAME}.pbip", {"version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
                                       "settings": {"enableAutoRecovery": True}})
    build_model(HERE)
    build_report(HERE, theme)
    print(f"Written {NAME}.pbip, {len(MEASURES)} measures")


if __name__ == "__main__":
    main()
