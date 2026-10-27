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

## Pages

### 1. Queue Overview (sheets `service_level_metrics_clean` + `abandoned_calls_clean`)
- Date range slicer.
- Cards: Offered, Answered, Abandoned, Service Level %, Abandon %, Calls Within SLA,
  Avg Speed of Answer, Avg Time to Abandon.
- Daily combo chart: Offered / Answered / Abandoned (columns), Service Level % and Abandon % (lines).
- Calls and abandoned calls by hour of day (from the call detail).
- Daily summary table. Service Level % turns red below target (80 %), and Abandon % turns red
  above target (5 %).
- List of abandoned calls with the caller's wait time before hanging up.

### 2. Agents (sheets `agent_status_clean` + `agent_performance_clean`)
- Date range slicer and agent drop-down.
- Cards: Agents Logged In, Logged In time, Allowed Status %, **Non-Allowed Status Time (red)**,
  Answered, Handled, AHT.
- Stacked bar per agent: hours in each status. Allowed statuses are blue/green, non-allowed
  statuses are shades of red, unknown status is grey.
- Donut: share of time in On Queue, Available, Break, Meeting, Training, Non-Allowed (red) and
  Unknown (grey).
- Status table, one row per agent, times as hh:mm:ss. Every non-allowed column gets a **red
  cell** when the agent spent any time in that status. The *Unknown (Other Off-Queue)* column
  turns grey when it has time. *Allowed %* is red below target (90 %).
- Performance table: Answered, Handled, AHT, Avg Talk / Hold / ACW, Held, Transferred.
- Bar chart: answered and handled calls by agent.

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

## Model

- Tables: the four sheets, **Calendar** (built from the dates in the workbook) and **Agents**
  (one row per agent_id). Calendar filters all four sheets through `report_date`. Agents filters
  both agent sheets.
- All ratios and averages are weighted (sum of numerator ÷ sum of denominator), so they stay
  correct over any date range:
  Service Level % = Calls Within SLA ÷ Answered. AHT = total handle time ÷ handled calls.
- Measures ending in "Time" or named ASA / AHT return hh:mm:ss text for display. Use the
  "Hours" / "(sec)" measures in charts and calculations.
