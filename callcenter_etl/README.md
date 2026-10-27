# Call-center ETL for Power BI (GL AUDI EHD)

`callcenter_etl.py` cleans four Genesys Cloud CSV exports and writes **one Excel workbook**
(`callcenter_powerbi_clean.xlsx`) with four sheets for Power BI:

| Sheet | Source file | One row = |
|---|---|---|
| `service_level_metrics_clean` | Queue performance (`AudiQPer-*.csv`) | queue × interval (currently 1 day) |
| `abandoned_calls_clean` | Conversation detail (`AudiAban-*.csv`) | one call (conversation) |
| `agent_status_clean` | Agent status (`AgentsSTA-*.csv`) | **one agent per day**: one column per status |
| `agent_performance_clean` | Agent performance (`AgentsPer-*.csv`) | **one agent per day**: daily totals and averages |

### The two agent sheets: each agent appears once per day

- **agent_status_clean**: for every status (Logged In, On Queue, Off Queue, Idle, Not Responding,
  Available, Away, Break, Meal, Busy, Meeting, Training) there are three columns:
  `<status>_seconds` (use in Power BI), `<status>_hhmmss` (readable, e.g. `08:00:50`) and
  `<status>_pct_of_logged_in` (share of the agent's logged-in time that day).
- **agent_performance_clean**: `answered`, `handled`, `held` and `transferred` are daily totals.
  `total_handle_seconds` is the daily handle time. `avg_handle_seconds` / `avg_handle_hhmmss`
  is the daily AHT (total handle time ÷ handled calls). Talk, hold and ACW averages also have
  seconds and hh:mm:ss columns.
- If a source file has several intervals per agent per day (e.g. an hourly export), they are
  combined into one daily row. Times and counts are added up, and averages are weighted by calls
  (talk and ACW by handled calls, hold by held calls), never averaged as averages.
  `intervals_in_day` shows how many source rows were combined, and `source_row` lists them.
  Exact duplicate rows are counted only once.
- A workbook from the earlier version (one row per status) is converted automatically
  the next time the script runs.

Each run also writes a data-quality report (`data_quality_report_<time>.md`), a table of
every check (`data_quality_checks_<time>.csv`), a log file, and a backup of the previous workbook.

## How to run

```bash
pip install pandas openpyxl
```

### 1. Set your folders (once)

Open `callcenter_etl.py` and edit the block near the top marked
**`>>> EDIT YOUR FOLDER PATHS HERE <<<`**:

```python
SERVICE_LEVEL_FOLDER = r"C:\CallCenter\input\queue_performance"     # AudiQPer-*.csv files
ABANDONED_FOLDER = r"C:\CallCenter\input\abandoned_calls"            # AudiAban-*.csv files
AGENT_STATUS_FOLDER = r"C:\CallCenter\input\agent_status"            # AgentsSTA-*.csv files
AGENT_PERFORMANCE_FOLDER = r"C:\CallCenter\input\agent_performance"  # AgentsPer-*.csv files

OUTPUT_FOLDER = r"C:\CallCenter\output"   # callcenter_powerbi_clean.xlsx goes here
```

Replace the paths with your real folders and keep the `r` in front of the quotes. Point the
Power BI `ExcelFilePath` parameter at `OUTPUT_FOLDER\callcenter_powerbi_clean.xlsx`.

### 2. Every day

Save each new export into its folder, then run:

```bash
python callcenter_etl.py
```

- Every CSV in the four folders that has **not been processed before** is processed, oldest
  first, and added to the workbook.
- Processed files are remembered in `OUTPUT_FOLDER\processed_files.json`, so running twice never
  loads a file twice. A file that is overwritten with new content is processed again.
- If you missed a few days, all waiting files are processed in one run.
- A folder with no new file keeps its sheet unchanged.
- A file is marked as processed only after the workbook was saved. If a run fails, the file is
  picked up again next time.

### Other options

```bash
python callcenter_etl.py --reprocess-all     # process every file in the folders again (no duplicates)
python callcenter_etl.py --no-history        # rebuild the workbook from scratch from all files
python callcenter_etl.py --agent-status-dir "D:\exports\status"   # override one folder for this run
python callcenter_etl.py --files-only --agent-status AgentsSTA-26Oct26.csv   # just one file, ignore folders
```

Other options: `--skill "GL AUDI EHD"`, `--skill-match token|exact`, `--drop-caller-number`
(remove the caller phone number column), `--output-dir`, `--output-name`. All other settings
(date formats, tolerances, blank handling) are in the `Config` class.

Exit codes: 0 = success, 2 = a folder was not found or a file could not be processed (its sheet is
kept unchanged and the reason is in the report), 1 = fatal error (e.g. the workbook is open in Excel).

## How appending works (past + new dates)

Each sheet has a **replacement key**. Rows from the new file replace existing rows with the
same key. All other existing rows are kept.

| Sheet | Key | Effect |
|---|---|---|
| service_level_metrics_clean | report_date + queue_name | Re-running the month's file updates the days it covers. No duplicates. |
| abandoned_calls_clean | conversation_id | Each call appears once. The newest export wins (calls that were still in progress get updated). |
| agent_status_clean | report_date + agent_id | Each daily file adds a day. Re-running the same day replaces that day. |
| agent_performance_clean | report_date + agent_id | Same as above. |

Safety rules:
- An incomplete interval (`Interval Complete = FALSE`) never replaces a completed one.
- The workbook is written to a temporary file first and then swapped in. A backup of the
  previous version is kept as `callcenter_powerbi_clean.backup.xlsx`.

## What the script found in the real files

1. **Daylight-saving shift (queue file).** From 25 Oct the daily intervals run
   `11:00 PM → 11:00 PM`. The interval starting `10/25 11:00 PM` holds the **26 Oct** calls:
   30 offered, the same as the 30 conversations dated 26 Oct in the call-detail file.
   `report_date` is therefore the date of the interval's **midpoint** for daily intervals.
   Use `report_date`, not `interval_start`, for all date analysis.
2. **Service Level definition.** `Service Level % = (Answer − Over SLA) / Answer` reproduces
   the source value exactly (19/28 = 67.86 %, 4/7 = 57.14 %). `Handled In SLA *` gives the same
   number only for completed intervals (it counts only finished calls: 1 instead of 4 on 27 Oct).
   The script therefore adds `answered_within_sla = answered − over_sla` and keeps
   `handled_in_sla` as the source value.
3. **Abandon %** in the source = Abandon / Offer (2/30 = 6.67 %). Abandoned calls are not in
   the Service Level denominator.
4. **Agent statuses overlap. Never add all statuses together.**
   - Logged In = On Queue + Off Queue (verified for every agent)
   - Idle and Not Responding happen *while On Queue*
   - Available, Away, Break, Meal, Busy, Meeting and Training are presences *while Off Queue*.
     For 22 agents, Off Queue time is larger than the exported presences. Some presence
     statuses are probably not in the export.
   - Only add together statuses from the same level, e.g. Break + Meal + Meeting, but never Logged In + Idle.
5. **Agent performance vs queue (WARNING).** On 26 Oct the agent file has **6** answered calls.
   The queue file has **28**. The 6 calls in the agent file are all afternoon calls (after
   ~13:15), so the agent export probably uses a different time zone or day boundary. See the
   questions below.
6. Blank count cells in Genesys mean 0. They are loaded as 0, and averages stay blank.
   Future days in the month file are labelled `data_status = "Incomplete interval (no activity yet)"`.
7. The skill filter keeps 32 of 80 agents in each agent file. 26 excluded agents have blank
   Skills. `GL AUDI VIP EHD` is not treated as a match.

## Duplicates

Duplicates are **flagged, never deleted automatically**:
- `exact_duplicate`: every column is identical.
- `duplicate_key_conflict`: same natural key (interval + queue, conversation_id, or interval +
  agent_id) but different values.
Both appear in the `dq_flags` column and in the report. Across runs, the replacement-key rule
above prevents the same data from being loaded twice.

## Columns worth knowing

- `dq_flags`: every row-level problem found for that row (blank = clean).
- `source_file`, `source_row`, `load_timestamp`: trace each row back to its CSV line.
- Durations are in **seconds** (`*_seconds`). The agent sheets also show them as readable
  `*_hhmmss` text. Original strings are kept in `time_to_abandon_raw` and `avg_handle_raw`.
- Percentages are **fractions** (0.6786). Set the Power BI format to Percentage.
- `not_responding_count` (agent performance): number of Not Responding events. Filled when the
  Agent Performance export includes the Genesys column "Alert - No Answer"; otherwise left blank
  (unknown), never 0.
- `*_source` = value from Genesys. `*_calc` = recalculated by the script for validation.
- `total_handle_seconds = avg_handle_seconds × handled` and
  `total_answer_wait_seconds = asa_seconds × answered`. These give correctly weighted averages.

## Power BI model

1. Import the four sheets. Set `report_date` to type **Date**.
2. Add a Date table (`CALENDARAUTO()` or your own) and relate `Date[Date]` to `report_date`
   on all four tables.
3. Optional: an Agent table (distinct `agent_id`, `agent_name`) related to both agent tables.
4. In visuals, filter `service_level_metrics_clean[data_status]` to exclude future days.

## Recommended KPI measures (DAX)

Ratios are always **SUM(numerator) / SUM(denominator)**. Never average a percentage or an
average column.

```DAX
-- Queue (service_level_metrics_clean = SL)
Offered Calls          = SUM ( SL[offered] )
Answered Calls         = SUM ( SL[answered] )
Abandoned Calls        = SUM ( SL[abandoned] )
Calls Within SLA       = SUM ( SL[answered_within_sla] )
Service Level %        = DIVIDE ( [Calls Within SLA], [Answered Calls] )
Abandon %              = DIVIDE ( [Abandoned Calls], [Offered Calls] )
Answer %               = DIVIDE ( [Answered Calls], [Offered Calls] )
ASA (sec)              = DIVIDE ( SUM ( SL[total_answer_wait_seconds] ), [Answered Calls] )
Queue AHT (sec)        = DIVIDE ( SUM ( SL[total_handle_seconds] ), SUM ( SL[handled] ) )

-- Call detail (abandoned_calls_clean = AB)
Call Records           = COUNTROWS ( AB )
Total Abandoned Calls  = SUM ( AB[abandoned_flag] )
Abandon % (detail)     = DIVIDE ( [Total Abandoned Calls], [Call Records] )
Avg Time to Abandon    = CALCULATE ( AVERAGE ( AB[time_to_abandon_seconds] ), AB[is_abandoned] = TRUE () )

-- Agent status (agent_status_clean = ST). One column per status; one row per agent per day.
Logged In Hours        = SUM ( ST[logged_in_seconds] ) / 3600
On Queue Hours         = SUM ( ST[on_queue_seconds] ) / 3600
Break Hours            = SUM ( ST[break_seconds] ) / 3600        -- same pattern for every status
On Queue % of Logged In = DIVIDE ( SUM ( ST[on_queue_seconds] ), SUM ( ST[logged_in_seconds] ) )
Idle % of On Queue     = DIVIDE ( SUM ( ST[idle_seconds] ), SUM ( ST[on_queue_seconds] ) )
Agents Logged In       = CALCULATE ( DISTINCTCOUNT ( ST[agent_id] ), ST[agent_logged_in] = TRUE () )

-- Agent performance (agent_performance_clean = AP)
Answered by Agent      = SUM ( AP[answered] )
Handled by Agent       = SUM ( AP[handled] )
AHT by Agent (sec)     = DIVIDE ( SUM ( AP[total_handle_seconds] ), [Handled by Agent] )
AHT (mm:ss)            = VAR s = [AHT by Agent (sec)]
                         RETURN IF ( NOT ISBLANK ( s ),
                                FORMAT ( INT ( s / 60 ), "00" ) & ":" & FORMAT ( MOD ( s, 60 ), "00" ) )
```

Status distribution visual: put the hour measures of one level in a stacked bar, e.g. On Queue +
Off Queue, or Available + Away + Break + Meal + Busy + Meeting + Training. For a status slicer,
unpivot the `*_seconds` columns in Power Query (Transform → Unpivot Columns). Do not use the
`*_pct_of_logged_in` columns in totals: they are per agent per day, so averaging them gives the wrong
result. Use the DIVIDE measures above instead.

## Open questions (the script uses the defaults shown)

1. **Agent file time zone.** Agent performance shows 6 answered calls on 26 Oct and the queue
   shows 28. Which time zone are the agent exports in, and is the interval really the full day?
   *Default: kept as exported and flagged.*
2. **Service Level.** Use `answered_within_sla` (matches Genesys `Service Level %`) or the
   `Handled In SLA *` column? What is the SLA threshold in seconds and the target %? The target
   column is blank. *Default: answered_within_sla.*
3. **Skill filter.** Should an agent with several skills that include `GL AUDI EHD` be kept
   (`token` mode)? Or only agents whose Skills cell is exactly `GL AUDI EHD` (`exact` mode)?
   *Default: token. Both give the same result for today's files.*
4. **Caller numbers.** The call-detail `Remote` column has customer phone numbers. Keep them in
   the Power BI file? *Default: kept. Use `--drop-caller-number` to remove them.*
5. **Interacting time.** Should the script derive On-Queue interacting time
   (On Queue − Idle − Not Responding)? *Default: not derived.*
6. **Queues.** The queue and call-detail files include both `GL AUDI EHD` and `GL AUDI VIP EHD`
   in their filter (call detail only contains GL AUDI EHD rows today). Should these files also be
   limited to GL AUDI EHD? *Default: no filter.*
