# report_tools.py
# Shared building blocks for the report scripts (top_customers.py and
# monday_summary.py). You don't run this file yourself; the reports use it.
#
# It does four jobs:
#   1. finds the cleaned spreadsheet made by clean.py
#   2. loads just the columns the reports need (never names, emails,
#      phones, companies, or notes)
#   3. works out which months each person was actually a customer
#   4. applies the privacy rule: no group smaller than 3 customers is shown

import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

CLEAN_FILE_NAME = "customers-clean.xlsx"
SMALLEST_GROUP = 3          # privacy rule: hide any group with fewer customers

# Month column headings look like "Oct 2025" (or "October 2025").
MONTH_HEADER_FORMATS = ["%b %Y", "%B %Y", "%Y-%m"]


# --- 1. Finding the files -------------------------------------------------

def find_files(argument):
    """Given the spreadsheet's location, return (cleaned file, reports folder).
    You can also point it straight at customers-clean.xlsx."""
    given = Path(argument).expanduser().resolve()
    if given.name == CLEAN_FILE_NAME:
        clean_file = given
    else:
        clean_file = given.parent / "reports" / CLEAN_FILE_NAME
    if not clean_file.exists():
        sys.exit(f"Couldn't find {clean_file}\n"
                 "Run clean.py first (or use run_all.py, which does everything).")
    # Warn if the spreadsheet was changed after it was last cleaned.
    if given != clean_file and given.exists() and \
            given.stat().st_mtime > clean_file.stat().st_mtime:
        print("Warning: the spreadsheet changed after it was cleaned. "
              "Run clean.py again (run_all.py does this for you).", file=sys.stderr)
    return clean_file, clean_file.parent


# --- 2. Loading the cleaned spreadsheet ------------------------------------

def month_start(header):
    """If a heading is a month like "Oct 2025", return Oct 1, 2025."""
    for layout in MONTH_HEADER_FORMATS:
        try:
            return datetime.strptime(str(header).strip(), layout).date()
        except ValueError:
            pass
    return None


def as_number(value):
    """Money cells should already be numbers after cleaning. Anything else
    (a cell clean.py couldn't fix) is treated as unknown."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return None


def load_customers(clean_file):
    """Read the cleaned file. Returns (months, customers).
    months: the last 12 month columns, oldest first, as dates.
    customers: one small record per customer with ONLY these details:
      id, plan, state, signup date, and 12 monthly amounts."""
    book = load_workbook(clean_file, read_only=True, data_only=True)
    rows = book.worksheets[0].iter_rows(values_only=True)
    headers = [str(h).strip() if h is not None else "" for h in next(rows)]

    def column(name):
        lowered = [h.lower() for h in headers]
        return lowered.index(name) if name in lowered else None

    id_col, plan_col = column("customer id"), column("plan")
    state_col, signup_col = column("state"), column("signup date")
    if id_col is None:
        sys.exit("The cleaned file has no 'Customer ID' column.")

    # Find the month columns and keep the most recent 12.
    month_cols = sorted((month_start(h), i) for i, h in enumerate(headers) if month_start(h))
    month_cols = month_cols[-12:]
    if not month_cols:
        sys.exit("The cleaned file has no month columns (like 'Oct 2025').")
    months = [m for m, _ in month_cols]

    customers, unreadable = [], 0
    for row in rows:
        if all(v is None for v in row):
            continue
        signup = row[signup_col] if signup_col is not None else None
        if isinstance(signup, datetime):
            signup = signup.date()
        amounts = []
        for _, i in month_cols:
            number = as_number(row[i])
            if number is None and row[i] is not None:
                unreadable += 1                 # text clean.py couldn't fix
            amounts.append(number)
        customers.append({
            "id": row[id_col] or "(no ID)",
            "plan": (row[plan_col] if plan_col is not None else None) or "(no plan)",
            "state": (row[state_col] if state_col is not None else None) or "",
            "signup": signup if isinstance(signup, date) else None,
            "amounts": amounts,
        })
    book.close()
    if unreadable:
        print(f"Note: {unreadable} money cell(s) weren't numbers and were skipped. "
              "See the 'Cleaning notes' tab.", file=sys.stderr)
    return months, customers


# --- 3. Which months was someone a customer? -------------------------------

def customer_months(customer):
    """The months (as positions 0-11) this person counted as a customer:
    from their first month on the books (signup, or the start of the 12
    months if they joined earlier) up to the last month they paid.
    Months after they stopped paying don't count. Returns [] if they
    never paid in these 12 months."""
    amounts = customer["amounts"]
    paid = [i for i, a in enumerate(amounts) if a and a > 0]
    if not paid:
        return []
    first = next(i for i, a in enumerate(amounts) if a is not None)
    return list(range(first, paid[-1] + 1))


def is_paying(customer, position):
    amount = customer["amounts"][position]
    return amount is not None and amount > 0


# --- 4. The privacy rule ----------------------------------------------------

def protect_groups(groups, other_label):
    """groups: {group name: [one item per customer]}.
    Any group with fewer than 3 customers is folded into an "Other" group.
    If "Other" would still be under 3, the next-smallest group joins it too,
    so nobody can work out a small group by subtracting from the total.
    Returns a list of (label, items). A group that can't reach 3 at all is
    left out entirely."""
    visible = {name: list(items) for name, items in groups.items() if items}
    other_items, other_names = [], []
    for name in [n for n, items in visible.items() if len(items) < SMALLEST_GROUP]:
        other_items += visible.pop(name)
        other_names.append(name)
    while other_names and len(other_items) < SMALLEST_GROUP and visible:
        smallest = min(visible, key=lambda n: len(visible[n]))
        other_items += visible.pop(smallest)
        other_names.append(smallest)
    result = list(visible.items())
    if other_names and len(other_items) >= SMALLEST_GROUP:
        result.append((f"{other_label} ({', '.join(sorted(other_names))})", other_items))
    return result


# --- Small formatting helpers ------------------------------------------------

def month_name(month):
    return month.strftime("%B %Y")        # "September 2026"


def dollars(amount):
    """$12,345 (whole dollars, for plain-English reports)."""
    sign = "-" if amount < 0 else ""
    return f"{sign}${abs(amount):,.0f}"
