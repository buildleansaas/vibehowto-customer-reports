# Customer reports project

## Privacy rules (always follow these)
- My real spreadsheet is in ~/Documents/work-data. Never open, read, list, or search anything there, and never run scripts on it. I run the scripts on the real file myself.
- Build and test everything using sample-customers.xlsx, which is fake.
- Reports must never include names, emails, phone numbers, companies, or notes. Refer to customers only by Customer ID.
- Summary reports show totals only. Hide any group with fewer than 3 customers.

## How scripts should work
- Python 3, using only built-in modules plus openpyxl.
- Each script takes the spreadsheet's location as its first argument, works with .xlsx or .csv, and saves its results in a "reports" folder next to that spreadsheet. Never change the original file.
- Add short comments explaining each part, written for someone who has never coded.
