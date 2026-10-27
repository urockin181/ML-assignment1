# Call Center Dashboard (Power BI project, PBIP)

`CallCenterDashboard_PBIP.zip` contains a Power BI project (PBIP) that reads the workbook created by
`../callcenter_etl.py` (`callcenter_powerbi_clean.xlsx`).

## Open it

1. Unzip `CallCenterDashboard_PBIP.zip` and keep the extracted items together: `CallCenterDashboard.pbip`, `CallCenterDashboard.Report/`
   and `CallCenterDashboard.SemanticModel/`.
2. Open `CallCenterDashboard.pbip` in **Power BI Desktop** (a version from mid-2025 or newer).
   If Desktop asks, enable *File → Options → Preview features → Power BI Project (.pbip) save
   option* and *Store reports using enhanced metadata format (PBIR)*. Newer versions have both on
   by default.
3. Go to *Home → Transform data → Edit parameters*. Set **ExcelFilePath** to the full path of
   `callcenter_powerbi_clean.xlsx`, e.g. `C:\CallCenter\output\callcenter_powerbi_clean.xlsx`.
4. Click **Refresh**. Each time `callcenter_etl.py` adds a new day, click Refresh again.

## Design

Audi-style theme (custom theme `AudiEHD.json` inside the project): black header bar with an
Audi-red line, light silver page, white rounded tiles. Black means on target, **Audi red means
alert**, and silver/grey means unknown. Each page fits one 16:9 screen with no scrolling.

## Pages

### 1. Queue Overview (sheets `service_level_metrics_clean` + `abandoned_calls_clean`)
- Header with a date range slicer.
- KPI tiles: Offered, Answered, Abandoned, Abandon % (red above 5 %), Service Level % (red
  below 80 %), Avg Speed of Answer.
- Daily calls (answered in black, abandoned stacked in red) with a Service Level % line.
- Gauge: Service Level against the 80 % target.
- Calls by hour of day (silver) with abandoned calls (red).
- Compact list of abandoned calls, newest first.

### 2. Agents (sheets `agent_status_clean` + `agent_performance_clean`)
- Header with date range and agent slicers.
- KPI tiles: Agents logged in, Breaks over 1 hour (red), Avg break per agent,
  Not Responding time (red), Not Responding events (red), On Queue % of logged-in time.
- **Break per agent**: bars turn red when an agent's break is over the 1-hour limit
  (with a multi-day date range: the average break per logged-in day).
- Break times are shown as **"45 min"** below one hour and as **"2h 06m"** from one hour up
  (KPI tile, chart labels and scorecard).
- **Not Responding per agent**: minutes in Not Responding; hovering shows the duration and the
  number of events. Only agents with Not Responding time are listed.
- **On Queue and Idle per agent** (hours); hovering shows On Queue % of logged-in time.
- **Other statuses per agent** (minutes): Meal, Away, Busy, Meeting, Training.
- **Agent scorecard**: Logged in, On Queue, On Queue %, Idle, Break (red cell over 1 hour),
  Not Responding (red cell when > 0), NR events, Meal (red cell when > 0), Answered, AHT.
- Agents with the skill who never logged in are hidden on this page.

**Not Responding events (count).** The current exports only contain the *duration* of Not
Responding. The number of events shows "n/a" until you add the Genesys column **"Alert - No Answer"**
to the Agent Performance export. `callcenter_etl.py` picks it up automatically
(`not_responding_count`) and the dashboard shows it after a refresh.

## Status rules

| Allowed (normal colors) | Not allowed (red) | Unknown (grey) |
|---|---|---|
| On Queue (includes Idle) | Away | Other Off-Queue (*) |
| Idle | Meal | |
| Available | Busy | |
| Break | Not Responding (missed alerts while On Queue) | |
| Meeting | | |
| Training | | |

(*) **Other Off-Queue** = Off Queue time not covered by any exported presence status (Available,
Away, Break, Meal, Busy, Meeting, Training). Genesys records it under a status that isn't in the
export, so it is shown in grey as "Unknown". It is **not** counted as non-allowed time.

Allowed time + Non-allowed time + Unknown time = Logged In time for every agent.
Allowed Status % = Allowed time ÷ Logged In time.

Idle is part of On Queue, so it is shown as its own column in the table but is not stacked a
second time in the bar chart.

## Targets (edit in the model)

The source data has no targets, so these measures hold placeholder values. Change them in
*Model view* or in `SemanticModel/definition/tables/*.tmdl`:

| Measure | Table | Value |
|---|---|---|
| Service Level Target % | service_level_metrics_clean | 0.80 |
| Abandon Target % | service_level_metrics_clean | 0.05 |
| Allowed Status Target % | agent_status_clean | 0.90 |
| Break Limit (min) | agent_status_clean | 60 |

## Model

- Tables: the four sheets, **Calendar** (built from the dates in the workbook) and **Agents**
  (one row per agent_id). Calendar filters all four sheets through `report_date`. Agents filters
  both agent sheets.
- All ratios and averages are weighted (sum of numerator ÷ sum of denominator), so they stay
  correct over any date range:
  Service Level % = Calls Within SLA ÷ Answered. AHT = total handle time ÷ handled calls.
- Measures ending in "Time" or named ASA / AHT return hh:mm:ss text for display. Use the
  "Hours" / "(sec)" measures in charts and calculations.
