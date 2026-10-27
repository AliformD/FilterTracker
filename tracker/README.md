# Filter tracker

One folder per job number (e.g. `2250823/`) holding `data.json` plus its spreadsheets:
- `jobnumber_PROJECT_FilterLog.xlsx` - the file for the PM: tab 1 Filter Log, tab 2 Filter Orders
- `jobnumber_PROJECT_OpenQuestions.xlsx` - just for us, not filed

Rebuild a job: `cd tracker/<job number> && python3 ../build.py`
Blank PM templates: `cd tracker && python3 build.py --template` (writes tracker/templates/)
