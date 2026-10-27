# Call-center ETL for Power BI (GL AUDI EHD)

`callcenter_etl.py` cleans four Genesys Cloud CSV exports and writes **one Excel workbook**
(`callcenter_powerbi_clean.xlsx`) with four sheets for Power BI:

| Sheet | Source file | One row = |
|---|---|---|
| `service_level_metrics_clean` | Queue performance (`AudiQPer-*.csv`) | queue × interval (currently 1 day) |
| `abandoned_calls_clean` | Conversation detail (`AudiAban-*.csv`) | one call (conversation) |
| `agent_status_clean` | Agent status (`AgentsSTA-*.csv`) | agent × interval × **status** (long format) |
| `agent_performance_clean` | Agent performance (`AgentsPer-*.csv`) | agent × interval |

Each run also writes a data-quality report (`data_quality_report_<time>.md`), a table of
every check (`data_quality_checks_<time>.csv`), a log file, and a backup of the previous workbook.

## How to run

```bash
pip install pandas openpyxl

# Monthly queue files + daily agent files
python callcenter_etl.py \
  --service-level AudiQPer-Oct26.csv \
  --abandoned AudiAban-Oct26.csv \
  --agent-status AgentsSTA-26Oct26.csv \
  --agent-performance AgentsPer-26Oct26.csv \
  --output-dir output

# Next day: only the new agent files. The queue sheets are kept as they are.
python callcenter_etl.py --agent-status AgentsSTA-27Oct26.csv \
  --agent-performance AgentsPer-27Oct26.csv --output-dir output
```

Options: `--skill "GL AUDI EHD"`, `--skill-match token|exact`, `--no-history` (rebuild the
workbook from scratch), `--drop-caller-number` (remove the caller phone number column),
`--output-name`. All other settings (date formats, tolerances, blank handling) are in the
`Config` class at the top of the script.

Exit codes: 0 = success, 2 = at least one file could not be processed (its sheet is kept
unchanged and the reason is in the report), 1 = fatal error (e.g. the workbook is open in Excel).

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
   - The `status_group` column holds these levels. Only sum or compare statuses inside one group.
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
- Durations are in **seconds** (`*_seconds`). Agent status also has `duration_minutes` and
  `duration_hours`. Original strings are kept in `time_to_abandon_raw` and `avg_handle_raw`.
- Percentages are **fractions** (0.6786). Set the Power BI format to Percentage.
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

-- Agent status (agent_status_clean = ST)
Status Hours           = SUM ( ST[duration_hours] )
Logged In Hours        = CALCULATE ( [Status Hours], ST[status] = "Logged In" )
% of Logged In Time    = DIVIDE ( [Status Hours],
                           CALCULATE ( [Status Hours], REMOVEFILTERS ( ST[status], ST[status_group], ST[status_sort] ),
                                       ST[status] = "Logged In" ) )
% Within Status Group  = DIVIDE ( [Status Hours],
                           CALCULATE ( [Status Hours], REMOVEFILTERS ( ST[status], ST[status_sort] ) ) )

-- Agent performance (agent_performance_clean = AP)
Answered by Agent      = SUM ( AP[answered] )
Handled by Agent       = SUM ( AP[handled] )
AHT by Agent (sec)     = DIVIDE ( SUM ( AP[total_handle_seconds] ), [Handled by Agent] )
AHT (mm:ss)            = VAR s = [AHT by Agent (sec)]
                         RETURN IF ( NOT ISBLANK ( s ),
                                FORMAT ( INT ( s / 60 ), "00" ) & ":" & FORMAT ( MOD ( s, 60 ), "00" ) )
```

Status distribution visual: use a stacked bar with **one** `status_group` as a filter
(e.g. "4 Presence (Off Queue)" or "2 Queue State") and the measure `% Within Status Group`.

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
