# run_all.py
# The one command to run every Monday. It cleans the spreadsheet, then
# makes both reports, then tells you which files are safe to share.
#
#   python3 run_all.py path/to/my-spreadsheet.xlsx
#
# The original spreadsheet is never changed. Everything is saved in a
# "reports" folder next to it.

import subprocess
import sys
from pathlib import Path

# The scripts to run, in order. Each one gets the spreadsheet's location.
STEPS = ["clean.py", "top_customers.py", "monday_summary.py"]


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python3 run_all.py path/to/my-spreadsheet.xlsx")
    spreadsheet = Path(sys.argv[1]).expanduser().resolve()
    if not spreadsheet.exists():
        sys.exit(f"File not found: {spreadsheet}")

    # The scripts live in the same folder as this file.
    script_folder = Path(__file__).resolve().parent

    for step in STEPS:
        print(f"\n=== Running {step} ===", flush=True)
        result = subprocess.run([sys.executable, str(script_folder / step), str(spreadsheet)])
        if result.returncode != 0:
            sys.exit(f"\n{step} didn't finish, so the steps after it were skipped.")

    reports = spreadsheet.parent / "reports"
    print(f"""
=== All done. Your files are in {reports} ===

SAFE TO SHARE with your boss:
  monday-summary.txt   totals only, no customer details, small groups hidden

INTERNAL ONLY (don't send outside the team):
  top-customers.csv    lists each customer's spending by Customer ID

KEEP PRIVATE:
  customers-clean.xlsx a full copy of the data, WITH names, emails, and phone numbers
""")


if __name__ == "__main__":
    main()
