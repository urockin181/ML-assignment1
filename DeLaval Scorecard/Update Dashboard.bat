@echo off
REM Prepares the scorecard data for the month found in "Source Files" (or pass --month YYYY-MM).
cd /d "%~dp0"
python scripts\prepare_scorecard_data.py %*
pause
