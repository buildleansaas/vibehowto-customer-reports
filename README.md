# Customer reports with AI (without sharing your data)

The finished code from the vibehowto guide **[Get reports from your customer spreadsheet with AI (without sharing it)](https://vibehowto.com/solve/spreadsheet)**.

Every script here was written by Claude Code, working only from a spreadsheet's column names and a fake sample file. It never saw real customer data. You run the scripts on your real spreadsheet yourself.

## What's here

| File | What it does |
|---|---|
| `CLAUDE.md` | The project's privacy rules, which Claude reads every session |
| `.claude/settings.json` | A lock that stops Claude from opening `~/Documents/work-data` |
| `clean.py` | Cleans the messy export into `reports/customers-clean.xlsx`, with a "Cleaning notes" tab |
| `top_customers.py` | Top customers by spend, with Customer IDs only |
| `monday_summary.py` | A plain-English summary for your boss, with totals only and small groups hidden |
| `report_tools.py` | Shared rules both reports use (you never run it yourself) |
| `run_all.py` | Cleans, then runs both reports, with one command |
| `make_sample_data.py`, `sample-customers.xlsx` | Fake practice data |

## Run it

```bash
python3 -m pip install --user openpyxl
python3 run_all.py ~/Documents/work-data/customers.xlsx
```

Reports appear in a `reports` folder next to your spreadsheet. Share `monday-summary.txt`. Keep `customers-clean.xlsx` to yourself, because it still has names, emails, and phone numbers.

All sample data is fake (example.com emails, 555-01xx phone numbers).
