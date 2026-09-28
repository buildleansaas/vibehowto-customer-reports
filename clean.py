# clean.py
# Makes a tidy copy of the customer spreadsheet. The original file is only
# read, never changed. The cleaned copy is saved as
#   reports/customers-clean.xlsx
# in the same folder as the spreadsheet you give it.
#
# What it fixes:
#   - extra spaces, blank rows, and duplicate customers (see Step 2 below)
#   - names as "First Last", emails in lowercase, phones as (415) 555-0101
#   - dates as 2026-03-14 ("187 days ago" counts back from today)
#   - states as two-letter codes, plans spelled the same way ("Biz" -> Business)
#   - money as plain numbers (blank or "-" after signup counts as 0)
# Anything it can't fix is listed on a "Cleaning notes" tab, by Customer ID only.
#
# To run it:  python3 clean.py path/to/customers.xlsx
# (A .csv file works too.)

import csv
import re
import sys
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


# --- Lookup tables --------------------------------------------------------

# Every US state (plus DC), so full names can be turned into two-letter codes.
STATES = {
    "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR",
    "CALIFORNIA": "CA", "COLORADO": "CO", "CONNECTICUT": "CT", "DELAWARE": "DE",
    "DISTRICT OF COLUMBIA": "DC", "FLORIDA": "FL", "GEORGIA": "GA",
    "HAWAII": "HI", "IDAHO": "ID", "ILLINOIS": "IL", "INDIANA": "IN",
    "IOWA": "IA", "KANSAS": "KS", "KENTUCKY": "KY", "LOUISIANA": "LA",
    "MAINE": "ME", "MARYLAND": "MD", "MASSACHUSETTS": "MA", "MICHIGAN": "MI",
    "MINNESOTA": "MN", "MISSISSIPPI": "MS", "MISSOURI": "MO", "MONTANA": "MT",
    "NEBRASKA": "NE", "NEVADA": "NV", "NEW HAMPSHIRE": "NH", "NEW JERSEY": "NJ",
    "NEW MEXICO": "NM", "NEW YORK": "NY", "NORTH CAROLINA": "NC",
    "NORTH DAKOTA": "ND", "OHIO": "OH", "OKLAHOMA": "OK", "OREGON": "OR",
    "PENNSYLVANIA": "PA", "RHODE ISLAND": "RI", "SOUTH CAROLINA": "SC",
    "SOUTH DAKOTA": "SD", "TENNESSEE": "TN", "TEXAS": "TX", "UTAH": "UT",
    "VERMONT": "VT", "VIRGINIA": "VA", "WASHINGTON": "WA",
    "WEST VIRGINIA": "WV", "WISCONSIN": "WI", "WYOMING": "WY",
}
STATE_CODES = set(STATES.values())

# The date layouts we know how to read, e.g. 03/14/2026, 14-Mar-2026,
# March 14, 2026. Numbers-only dates are read US-style (month first).
DATE_FORMATS = [
    "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y",
    "%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y",
    "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %B %Y", "%Y/%m/%d",
]

# Month column headings like "Oct 2025" or "October 2025".
MONTH_HEADER_FORMATS = ["%b %Y", "%B %Y", "%Y-%m"]

# A simple check for a real-looking email: something@something.something
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$")

# Our real plans. Any other plan name gets flagged on the notes tab.
PLAN_NAMES = ["Starter", "Pro", "Business"]
# Short forms people type, and the plan they mean (write these in lowercase).
PLAN_NICKNAMES = {"biz": "Business"}

# "187 days ago" is counted back from the day the script is run.
TODAY = date.today()


# --- Small helpers --------------------------------------------------------

def tidy_text(value):
    """Trim spaces from the ends and squeeze double spaces into one.
    Empty cells come back as None."""
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def is_blank(value):
    return tidy_text(value) is None


# --- One cleaner for each kind of column ----------------------------------
# Each cleaner returns (clean value, problem). "problem" is a short message
# for the notes tab, or None if all is well. Messages never repeat the
# original value, so no personal details end up on the notes tab.

def clean_id(value):
    text = tidy_text(value)
    if text is None:
        return None, "Missing Customer ID"
    # IDs should have no spaces inside and use capital letters (c-1037 -> C-1037).
    return re.sub(r"\s", "", text).upper(), None


def fix_word_case(word):
    """Capitalize a word only if it's ALL CAPS or all lowercase.
    Names already in mixed case (like McDonald) are left alone.
    Handles hyphens and apostrophes: o'brien -> O'Brien, smith-jones -> Smith-Jones."""
    if word != word.upper() and word != word.lower():
        return word
    return re.sub(r"[^\W\d_]+", lambda m: m.group(0).capitalize(), word.lower())


def clean_name(value):
    text = tidy_text(value)
    if text is None:
        return None, "Name is blank"
    problem = None
    # "Last, First" becomes "First Last".
    parts = [p.strip() for p in text.split(",")]
    if len(parts) == 2 and parts[0] and parts[1]:
        text = f"{parts[1]} {parts[0]}"
    elif len(parts) > 2:
        problem = "Name has more than one comma; order not changed"
    words = [fix_word_case(w) for w in text.split(" ")]
    return " ".join(words), problem


def clean_email(value):
    text = tidy_text(value)
    if text is None:
        return None, None
    text = text.replace(" ", "").lower()
    if text.startswith("mailto:"):
        text = text[len("mailto:"):]
    if not EMAIL_PATTERN.match(text):
        return None, "Email wasn't a real email address; left blank"
    return text, None


def clean_phone(value):
    text = tidy_text(value)
    if text is None:
        return None, None
    # Split off an extension like "x23" or "ext. 23", if there is one.
    main, extension = text, ""
    match = re.search(r"\s*(?:x|ext\.?|extension)\s*(\d+)\s*$", text, re.I)
    if match:
        main, extension = text[:match.start()], f" x{match.group(1)}"
    digits = re.sub(r"\D", "", main)
    # Drop the US country code "1" from the front of 11-digit numbers.
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}{extension}", None
    if len(digits) == 7:
        return text, "Phone has no area code; left as it was"
    return text, "Phone isn't a 10-digit US number; left as it was"


def parse_date(value):
    """Turn any date we recognise into a real date. Returns None if we can't."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    # A plain number in a date column is how Excel stores dates internally
    # (days since Dec 30, 1899). Only accept sensible years (about 1955-2119).
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 <= value <= 80000:
            return date(1899, 12, 30) + timedelta(days=int(value))
        return None
    text = tidy_text(value)
    if text is None:
        return None
    text = text.rstrip(".")
    # Relative dates: "187 days ago", "1 day ago", "today", "yesterday".
    match = re.fullmatch(r"(\d+)\s+days?\s+ago", text, re.I)
    if match:
        return TODAY - timedelta(days=int(match.group(1)))
    if text.lower() in ("today", "yesterday"):
        return TODAY - timedelta(days=0 if text.lower() == "today" else 1)
    # The same Excel date number, but saved as text (common in CSV files).
    if re.fullmatch(r"\d{5}(\.\d+)?", text):
        return parse_date(float(text))
    for layout in DATE_FORMATS:
        try:
            return datetime.strptime(text, layout).date()
        except ValueError:
            pass
    # Some exports add a time, e.g. "2026-03-14 09:30:00". Keep just the date.
    match = re.match(r"^(\d{4}-\d{2}-\d{2})[ T]", text)
    if match:
        return datetime.strptime(match.group(1), "%Y-%m-%d").date()
    return None


def clean_date(value):
    if is_blank(value):
        return None, None
    found = parse_date(value)
    if found is None:
        return tidy_text(value), "Date not recognised; left as it was"
    return found, None


def clean_state(value):
    text = tidy_text(value)
    if text is None:
        return None, None
    key = text.replace(".", "").upper()      # "N.Y." -> "NY"
    if key in STATE_CODES:
        return key, None
    if key in STATES:
        return STATES[key], None
    return text, "State not recognised; left as it was"


def clean_plan(value):
    text = tidy_text(value)
    if text is None:
        return None, f"Plan is blank (should be {', '.join(PLAN_NAMES)})"
    key = text.lower()
    # Short forms like "Biz" become the full plan name.
    if key in PLAN_NICKNAMES:
        return PLAN_NICKNAMES[key], None
    # "PRO ", "pro", " Pro" all become "Pro".
    for plan in PLAN_NAMES:
        if key == plan.lower():
            return plan, None
    # Anything else is tidied up but flagged, since it isn't one of our plans.
    tidied = " ".join(w.capitalize() for w in text.split(" "))
    return tidied, f"Plan isn't {', '.join(PLAN_NAMES[:-1])}, or {PLAN_NAMES[-1]}; left as it was"


def parse_money(value):
    """Turn things like "$1,200.00", "USD 99", " $49 ", "1,200 USD" into a number.
    Returns ("blank", None) for empty cells and dashes, or (None, None) if
    it doesn't look like money at all."""
    if isinstance(value, bool):
        return None, None
    if isinstance(value, (int, float)):
        return float(value), None
    text = tidy_text(value)
    if text is None or text in {"-", "--", "—", "–"}:
        return "blank", None
    negative = text.startswith("(") and text.endswith(")")   # (49.00) means -49
    text = re.sub(r"(?i)usd|\$|,|\s|\(|\)", "", text)
    if text.startswith("-"):
        negative, text = True, text[1:]
    if not re.fullmatch(r"\d+(\.\d+)?", text):
        return None, None
    amount = float(text)
    return (-amount if negative else amount), None


def clean_whole_number(value):
    """For count columns like Support Tickets: " 7 " -> 7."""
    if is_blank(value):
        return None, None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            return int(value), None
    text = tidy_text(value)
    if text and re.fullmatch(r"\d+(\.0+)?", text):
        return int(float(text)), None
    return tidy_text(value), "Not a whole number; left as it was"


# --- Reading the spreadsheet ----------------------------------------------

def read_rows(path):
    """Return every row of the file as a list of values (first sheet only)."""
    if path.suffix.lower() == ".csv":
        # "utf-8-sig" copes with the invisible marker Excel adds to CSV files.
        with open(path, newline="", encoding="utf-8-sig") as f:
            return [row for row in csv.reader(f)], "Customers"
    # read_only: just look, don't touch. data_only: use formula results.
    book = load_workbook(path, read_only=True, data_only=True)
    sheet = book.worksheets[0]
    rows = [list(r) for r in sheet.iter_rows(values_only=True)]
    title = sheet.title
    book.close()
    return rows, title


def find_header_row(rows):
    """The header is the first row (in the top 10) that has a Customer ID column."""
    for i, row in enumerate(rows[:10]):
        labels = [(tidy_text(v) or "").lower() for v in row]
        if "customer id" in labels:
            return i
    sys.exit("Couldn't find a 'Customer ID' column in the first 10 rows.")


def month_start(header):
    """If a column heading is a month like "Oct 2025", return Oct 1, 2025."""
    for layout in MONTH_HEADER_FORMATS:
        try:
            return datetime.strptime(header, layout).date()
        except ValueError:
            pass
    return None


def column_kind(header):
    """Decide what kind of data a column holds, from its heading."""
    h = header.lower()
    if h == "customer id":
        return "id"
    if h == "name":
        return "name"
    if "email" in h:
        return "email"
    if "phone" in h:
        return "phone"
    if h == "state":
        return "state"
    if h == "plan":
        return "plan"
    if month_start(header):
        return "money"
    if "date" in h or "login" in h:
        return "date"
    if "ticket" in h:
        return "count"
    return "text"


CLEANERS = {
    "id": clean_id, "name": clean_name, "email": clean_email,
    "phone": clean_phone, "state": clean_state, "plan": clean_plan,
    "date": clean_date, "count": clean_whole_number,
}


# --- The main job ---------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python3 clean.py path/to/customers.xlsx")
    source = Path(sys.argv[1]).expanduser().resolve()
    if not source.exists():
        sys.exit(f"File not found: {source}")

    rows, sheet_title = read_rows(source)
    header_index = find_header_row(rows)
    headers = [tidy_text(h) or f"Column {i + 1}"
               for i, h in enumerate(rows[header_index])]
    kinds = [column_kind(h) for h in headers]
    id_col = kinds.index("id")
    email_col = kinds.index("email") if "email" in kinds else None
    signup_col = next((i for i, h in enumerate(headers)
                       if h.lower() == "signup date"), None)

    fixed = Counter()       # how many cells of each kind were changed
    cleaned_rows = []       # each row's label, cleaned values, and notes
    blank_rows = 0

    # Step 1: clean every row, one cell at a time.
    for offset, raw in enumerate(rows[header_index + 1:]):
        sheet_row = header_index + 2 + offset       # row number as Excel shows it
        raw = list(raw) + [None] * (len(headers) - len(raw))
        if all(is_blank(v) for v in raw):
            blank_rows += 1
            continue

        clean = [None] * len(headers)
        row_notes = []
        for col, kind in enumerate(kinds):
            if kind == "money":
                continue                            # done below, needs signup date
            cleaner = CLEANERS.get(kind, lambda v: (tidy_text(v), None))
            clean[col], problem = cleaner(raw[col])
            if problem:
                row_notes.append((headers[col], problem))
            # Count it as a fix only if the value really changed
            # (a cell that was already a proper date doesn't count).
            already_date = isinstance(raw[col], datetime) and clean[col] == raw[col].date()
            if clean[col] != raw[col] and not already_date:
                fixed[kind] += 1

        # Money: blank or "-" counts as 0 once they've signed up, and stays
        # blank for months before they signed up.
        signup = clean[signup_col] if signup_col is not None else None
        signup = signup if isinstance(signup, date) else None
        blanks_left_unsure = []
        for col, kind in enumerate(kinds):
            if kind != "money":
                continue
            amount, _ = parse_money(raw[col])
            month = month_start(headers[col])
            if amount is None:
                clean[col] = tidy_text(raw[col])
                row_notes.append((headers[col], "Amount isn't a number; left as it was"))
            elif amount == "blank":
                if signup is None:
                    clean[col] = None
                    blanks_left_unsure.append(headers[col])
                elif (month.year, month.month) >= (signup.year, signup.month):
                    clean[col] = 0.0
                else:
                    clean[col] = None               # before signup: stays blank
            else:
                clean[col] = round(amount, 2)
                if signup and (month.year, month.month) < (signup.year, signup.month) and amount != 0:
                    row_notes.append((headers[col], "Amount charged before the signup date"))
            if clean[col] != raw[col]:
                fixed["money"] += 1
        if blanks_left_unsure and signup_col is not None:
            row_notes.append((", ".join(blanks_left_unsure),
                              "No readable signup date, so blank months were left blank (not 0)"))

        label = clean[id_col] or f"(no ID, row {sheet_row})"
        cleaned_rows.append({"label": label, "values": clean, "notes": row_notes})

    # Step 2: remove duplicates. The first copy is kept. If the later copy
    # has something the first copy is missing, that gap gets filled in.
    # A row counts as a duplicate if it has the same Customer ID, or the same
    # email and no ID of its own. Two rows with the same email but DIFFERENT
    # IDs are both kept and flagged, so a real customer is never deleted.
    kept, by_id, by_email = [], {}, {}
    duplicates_removed = 0
    for entry in cleaned_rows:
        clean = entry["values"]
        cid = clean[id_col]
        email = clean[email_col] if email_col is not None else None
        match = by_id.get(cid) if cid else None
        email_match = by_email.get(email) if email else None
        if match is None and email_match is not None:
            other_id = email_match["values"][id_col]
            if cid and other_id and other_id != cid:
                entry["notes"].append(("Email", f"Same email as {other_id} but a different "
                                       "ID; both kept. Check if they're the same customer."))
                email_match = None
            match = email_match
        if match is None:
            kept.append(entry)
            if cid:
                by_id[cid] = entry
            if email and email not in by_email:
                by_email[email] = entry
            continue

        # This row is a duplicate: fold it into the copy we're keeping.
        # Its own cleaning notes are dropped, since the row itself is gone.
        duplicates_removed += 1
        if cid:
            match["notes"].append(("", "Removed a duplicate copy of this customer"))
        else:
            match["notes"].append(("", f"Removed {entry['label']}: same email as this customer"))
        differing = []
        first = match["values"]
        for col in range(len(headers)):
            if is_blank(first[col]) and not is_blank(clean[col]):
                first[col] = clean[col]             # fill the gap from the copy
            elif not is_blank(clean[col]) and clean[col] != first[col]:
                differing.append(headers[col])
        if differing:
            match["notes"].append((", ".join(differing),
                                   "Duplicate copy disagreed here; kept the first copy's values"))

    # Gather the notes for the customers we kept, in spreadsheet order.
    notes = [(entry["label"], column, problem)
             for entry in kept for column, problem in entry["notes"]]

    # Step 3: save the cleaned copy plus the notes tab.
    out_folder = source.parent / "reports"
    out_folder.mkdir(exist_ok=True)
    out_file = out_folder / "customers-clean.xlsx"

    book = Workbook()
    sheet = book.active
    sheet.title = sheet_title
    sheet.append(headers)
    for entry in kept:
        sheet.append(entry["values"])
    # Make dates show as 2026-03-14 and money as plain numbers like 49.00.
    for row in sheet.iter_rows(min_row=2):
        for cell, kind in zip(row, kinds):
            if isinstance(cell.value, date):
                cell.number_format = "yyyy-mm-dd"
            elif kind == "money" and isinstance(cell.value, float):
                cell.number_format = "0.00"
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"                       # keep headings visible when scrolling

    note_sheet = book.create_sheet("Cleaning notes")
    note_sheet.append(["Customer ID", "Column", "What needs checking"])
    for cell in note_sheet[1]:
        cell.font = Font(bold=True)
    for note in notes:
        note_sheet.append(list(note))

    # Widen columns a little so they're readable.
    for ws in (sheet, note_sheet):
        for column in ws.columns:
            longest = max(len(str(c.value)) if c.value is not None else 0 for c in column)
            ws.column_dimensions[column[0].column_letter].width = min(max(10, longest + 2), 60)

    book.save(out_file)

    # Step 4: print a short summary (counts only, no customer details).
    print(f"Saved {out_file}")
    print(f"  Customers kept:       {len(kept)}")
    print(f"  Duplicates removed:   {duplicates_removed}")
    print(f"  Blank rows removed:   {blank_rows}")
    for kind in ["id", "name", "email", "phone", "state", "plan", "date", "money", "count", "text"]:
        if fixed[kind]:
            print(f"  {kind.title() + ' cells changed:':<22}{fixed[kind]}")
    print(f"  Items to check:       {len(notes)} (see the 'Cleaning notes' tab)")


if __name__ == "__main__":
    main()
