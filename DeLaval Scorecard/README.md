# DeLaval Scorecard (Power BI project)

This is a rebuild of `Original/DeLaval Scorecard DB.pbix` as a Power BI project (`.pbip`).
The report pages, visuals and DAX measures are the same as in the original. What changed is
how the data gets in: a Python script now reads all the monthly exports, cleans them and
writes **one** workbook (`Data/Scorecard_Data.xlsx`). Power BI reads only that workbook.

## Folder layout

| Path | What it is |
|---|---|
| `DeLaval Scorecard DB.pbip` | Open this in Power BI Desktop |
| `DeLaval Scorecard DB.Report/` | Report pages and visuals (same as the original) |
| `DeLaval Scorecard DB.SemanticModel/` | Tables, relationships and measures (TMDL) |
| `Source Files/` | Put the monthly exports here |
| `Data/Scorecard_Data.xlsx` | Created by the script; this is what Power BI loads |
| `scripts/prepare_scorecard_data.py` | Turns the source files into `Scorecard_Data.xlsx` |
| `Update Dashboard.bat` | Double-click on Windows to run the script |
| `config/agent_aliases.csv` | Spellings of agent names, mapped to the name shown on the dashboard |
| `config/settings.json` | Rules that used to be hard-coded in Power Query (thresholds, exclusions) |
| `Original/` | The original `.pbix`, kept for reference |

## One-time setup

1. Install Python 3.10 or newer from python.org. During setup, tick **Add python.exe to PATH**.
2. Open a command prompt in this folder and run: `pip install -r requirements.txt`
3. Open `scripts/prepare_scorecard_data.py` and set the four paths in the **PATHS** block at the
   top of the file (each is marked `# <-- CHANGE`):

   | Setting | What it points to |
   |---|---|
   | `DEFAULT_SOURCE` | the folder with the monthly source files |
   | `DEFAULT_OUTPUT` | the `Scorecard_Data.xlsx` file the script creates (Power BI reads it) |
   | `CONFIG_DIR` | the `config` folder (`agent_aliases.csv`, `settings.json`) |
   | `MODEL_EXPRESSIONS` | `DeLaval Scorecard DB.SemanticModel\definition\expressions.tmdl` (or `None`) |

   They are preset to `C:\DeLaval Scorecard\...`. If you put the folder exactly there, you
   don't need to change anything.
4. On older Power BI Desktop versions, turn on **File > Options and settings > Options >
   Preview features > Power BI Project (.pbip) save option** and **Store semantic model using
   TMDL format**. Current versions don't need this.

## Every month (example: in October, report on September)

1. Empty `Source Files/` and copy in the new exports for the month:

   | Export | How it is recognised |
   |---|---|
   | Genesys Calls (`.csv`) | has columns *Conversation ID*, *Total Queue* |
   | Genesys Status (`.csv`) | has columns *Agent Name*, *Primary Status* |
   | CSAT (`.csv`) | has columns *instance*, *string_value* |
   | Chat closed complete (`.csv`) | has columns *number*, *opened_for*, *assigned_to* |
   | Chat timed out (`.csv`) | has columns *work_item*, *rejected_by* |
   | Tasks Resolved by SD (`.csv`) | has columns *request_item*, *reassignment_count* |
   | Incidents Resolved by SD (`.csv`) | has columns *caller_id*, *reassignment_count* |
   | `<Month> ROTA.xlsx` | has sheets *L1*, *L2*, *BO* |
   | `<Month>.xlsx` (office attendance) | has the *DailyCheck* table |
   | `<Month> <Year> Team FCR.xlsx` | has sheet *L2 Tasks* (and *All Team*) |
   | Quality (`.xlsx`) | has sheet *Merged* |
   | Incident Management (`.xlsx`) | one sheet per month, e.g. *Sep*, *Oct* |
   | Knowledge Management (`.xlsx`) | has column *On-Time Delivery Rate* |

   File names don't matter; the script recognises each file from its columns and sheets.
   Keep **one** file of each kind. Exports that cover several months (Genesys, CSAT,
   Quality, Tasks and Incidents) are fine, because the script keeps only the report month.
2. Double-click **`Update Dashboard.bat`** (or run `python scripts/prepare_scorecard_data.py`).
   The dashboard always shows **one month**, taken from the source files: the latest month
   that most of the files reach. If the exports end on 30 September, it shows September.
   One stray record from the next month (for example a CSAT survey on 1 October) doesn't
   change the month; the script prints a warning instead. To force a month, run
   `python scripts/prepare_scorecard_data.py --month 2026-09`.
3. Read the messages it prints. Lines marked `WARNING` mean a source file is missing or
   incomplete. The same messages are saved in the **Log** sheet of `Data/Scorecard_Data.xlsx`.
4. Open `DeLaval Scorecard DB.pbip` and click **Refresh**. Save.

The script also points the Power BI parameter **DataFile** at `Data/Scorecard_Data.xlsx` on
your computer. To keep the workbook somewhere else (for example SharePoint or OneDrive), change
it under **Transform data > Edit parameters > DataFile**.

## Who appears where

Agents and their roles come from the monthly ROTA (sheets L1, L2 and BO, column *Grade*):

| ROTA grade | Dashboard section |
|---|---|
| L1 | L1 page |
| L2 | L2 & BO page, *L2 Service Desk* table |
| IM, IM CO | L2 & BO page, *Incident Management* table |
| KM | L2 & BO page, *Knowledge Management* table |
| PM | L2 & BO page, *Procurement Management* table |
| anything else (e.g. TM) | not shown |

So when someone joins, leaves or changes team, update the ROTA. Nothing needs to change in
Power BI. You can change the grade-to-section mapping in `config/settings.json` (`grade_to_role`).

The same person is often spelled differently across systems (`Mohamed Halim` in the ROTA,
`Mohamed Abdelhalim` in Genesys, `Amr AbdelwahabEXT` in ServiceNow). The script removes
`EXT`, ignores upper/lower case and uses `config/agent_aliases.csv` for the rest. If the
script lists a name under *"Names in the data that are not in the ROTA"* and it is really one
of your agents, add a line to that file.

## Differences from the original .pbix

The visuals, layout, formatting and every DAX measure are the same as the original. These
are the changes:

- **Month selection.** The old queries mixed fixed months (`Date.Month = 7`, sheet `Oct`,
  file `July.xlsx`) with "previous month". Every table now uses the same report month.
- **Agent filters.** The L1 page and the L2 & BO tables filtered on lists of names typed into
  the report (e.g. Procurement = *Mahmoud Hadhood*). They now filter on the role from the ROTA.
- **Name matching.** CSAT, chat, quality and attendance data were lost for some agents because
  of `EXT` suffixes and spelling differences (e.g. Habiba's chat timeouts, Sarah's CSAT). All
  names are now matched to the ROTA name.
- **Chat timeouts.** The old query silently dropped the first row of the file. It no longer does.
- **IM rates.** IM rates were stored as whole numbers, so 95 % became 100 %. They now keep
  decimals.
- **Removed relationships.** The Genesys ↔ Genesys_Abandoned link on *Total Queue Seconds*
  broke refresh whenever two calls had the same queue time. That is why the old query excluded
  4.112 s, 4.715 s, etc. It was inactive and unused, so it was removed. The unused
  Adherence ↔ Schedule many-to-many link was also removed.
- **FCR Team.** The *FCR Team* measure (hidden page) referred to a table that no longer
  existed. It is restored from the *All Team* sheet of the Team FCR workbook.
- **Removed queries.** The SharePoint and `C:\Users\...` paths are gone, and so is the unused
  "Errors in Schedule" query.

The hidden pages (Page 1–4, L1 Chat) were copied as they were.

## Troubleshooting

- **"more than one … file"**: two files of the same kind are in `Source Files/`. Remove the old one.
- **Power BI can't find the file**: run the script once on this computer, or set the
  `DataFile` parameter yourself.
- **A table on the dashboard is empty**: check the Log sheet. Usually the export for that part
  doesn't contain the report month.
- **Incident Management shows another month's figures**: the workbook has no sheet for the
  report month. The script then uses the sheet that best matches the current IM team and
  prints a warning. Add a sheet named after the month (e.g. `Sep`).
