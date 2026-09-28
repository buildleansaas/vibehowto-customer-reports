# make_sample_data.py
# Builds sample-customers.xlsx: 40 FAKE customers for building and testing
# report scripts. Nothing here is real. Emails use example.com and phone
# numbers use the 555-01xx range, which are reserved for fiction.
#
# The data is deliberately messy, like a real export: extra spaces, mixed
# capitalization, "Last, First" names, different phone/date/money formats,
# "Biz" for the Business plan, a misspelled plan, logins like "187 days ago",
# one duplicated customer, and one blank row.
#
# To run it:  python3 make_sample_data.py

import random
from datetime import date, datetime, timedelta

from openpyxl import Workbook

# --- Settings -------------------------------------------------------------

# A fixed "seed" makes the random choices the same every time, so the
# sample file comes out identical whenever you rebuild it.
random.seed(2026)

OUTPUT_FILE = "sample-customers.xlsx"

MONTHS = ["Oct 2025", "Nov 2025", "Dec 2025", "Jan 2026", "Feb 2026", "Mar 2026",
          "Apr 2026", "May 2026", "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026"]

# The first day of each month column, used to decide when someone joined.
MONTH_STARTS = [date(2025, 10, 1), date(2025, 11, 1), date(2025, 12, 1),
                date(2026, 1, 1), date(2026, 2, 1), date(2026, 3, 1),
                date(2026, 4, 1), date(2026, 5, 1), date(2026, 6, 1),
                date(2026, 7, 1), date(2026, 8, 1), date(2026, 9, 1)]

HEADERS = (["Customer ID", "Name", "Email", "Phone", "Company", "State", "Plan",
            "Signup Date"] + MONTHS +
           ["Last Login", "Support Tickets (90 days)", "Notes"])

FIRST_NAMES = ["Avery", "Jordan", "Riley", "Morgan", "Casey", "Taylor", "Jamie",
               "Quinn", "Parker", "Rowan", "Skyler", "Dana", "Emerson", "Hayden",
               "Kendall", "Logan", "Marley", "Noel", "Peyton", "Reese"]
LAST_NAMES = ["Sample", "Testerson", "Placeholder", "Example", "Demo", "Fakely",
              "Mockwell", "Dummy", "Notreal", "Pretend", "Sandbox", "Specimen"]
COMPANY_WORDS = ["Acme", "Globex", "Initech", "Umbrella", "Hooli", "Vandelay",
                 "Stark", "Wayne", "Wonka", "Tyrell", "Cyberdyne", "Soylent"]
COMPANY_ENDINGS = ["LLC", "Inc.", "Co", "Group", "Labs", "& Sons", "Partners"]

# States are written several ways on purpose ("CA", "ca ", "California").
# A few states appear only once or twice, to test "hide groups under 3".
STATE_SPELLINGS = {
    "CA": ["CA", "ca", " CA ", "California"],
    "TX": ["TX", "tx", "Texas", "TX "],
    "NY": ["NY", "ny", "New York"],
    "FL": ["FL", "Florida", "fl"],
    "IL": ["IL", "Illinois"],
    "WA": ["WA"],
    "VT": ["VT"],
}
STATE_WEIGHTS = {"CA": 10, "TX": 8, "NY": 7, "FL": 6, "IL": 5, "WA": 2, "VT": 1}

# Plans and their normal monthly price. Plan names get messy capitalization.
PLAN_PRICES = {"Starter": 49, "Pro": 199, "Business": 1200}
PLAN_SPELLINGS = {
    "Starter": ["Starter", "starter", "STARTER", " Starter"],
    "Pro": ["Pro", "pro", "PRO ", "Pro"],
    "Business": ["Business", "business", "BUSINESS", "Business "],
}

NOTES = ["", "", "", "Asked about annual billing", "Renewal call scheduled",
         "Unhappy with recent outage", "Champion left the company",
         "Wants SSO", "Referred by another customer", "Called re: invoice",
         "Considering downgrade", "Very happy - possible case study"]

TODAY = date(2026, 9, 26)


# --- Helpers that make values look messy --------------------------------

def messy_money(amount):
    """Turn a number into one of the many ways money shows up in exports."""
    if amount is None:
        return None                       # empty cell = no charge that month
    if amount == 0:
        return random.choice([0, "$0.00", "0", None, "-"])
    style = random.randint(1, 7)
    if style == 1:
        return f"${amount:,.2f}"          # $1,200.00
    if style == 2:
        return f"USD {amount:g}"          # USD 99
    if style == 3:
        return float(amount)              # a real number cell
    if style == 4:
        return f" ${amount:g} "           # extra spaces
    if style == 5:
        return f"{amount:.2f}"            # 99.00 stored as text
    if style == 6:
        return f"{amount:,.0f} USD"       # 1,200 USD
    return int(round(amount))             # plain whole number


def messy_date(d):
    """Show a date in one of several formats (or as a real date cell)."""
    style = random.randint(1, 6)
    if style == 1:
        return datetime(d.year, d.month, d.day)   # real Excel date
    if style == 2:
        return d.strftime("%m/%d/%Y")             # 03/15/2024
    if style == 3:
        return f"{d.month}/{d.day}/{d.strftime('%y')}"  # 3/15/24
    if style == 4:
        return d.strftime("%Y-%m-%d")             # 2024-03-15
    if style == 5:
        return d.strftime("%B %d, %Y")            # March 15, 2024
    return d.strftime("%d-%b-%Y")                 # 15-Mar-2024


def messy_phone(n):
    """Fake phone number in the 555-01xx range, in a random format."""
    area = random.choice(["312", "415", "212", "512", "305"])
    last = f"01{n:02d}"
    return random.choice([
        f"({area}) 555-{last}",
        f"{area}-555-{last}",
        f"{area}.555.{last}",
        f"+1 {area} 555 {last}",
        f"{area}555{last}",
        f"555-{last}",
    ])


def messy_name(first, last):
    """Names written a few different ways."""
    return random.choice([
        f"{first} {last}",
        f"{last}, {first}",
        f"{last.upper()}, {first.upper()}",
        f"{first.lower()} {last.lower()}",
        f"  {first}  {last} ",
    ])


def messy_email(first, last):
    email = f"{first.lower()}.{last.lower()}@example.com"
    return random.choice([email, email, email.upper(), f" {email}",
                          email.capitalize()])


def messy_company():
    name = f"{random.choice(COMPANY_WORDS)} {random.choice(COMPANY_ENDINGS)}"
    return random.choice([name, name, name.lower(), f" {name}  ", name.upper()])


def pick_state():
    code = random.choices(list(STATE_WEIGHTS), weights=STATE_WEIGHTS.values())[0]
    return random.choice(STATE_SPELLINGS[code])


# --- Spending patterns ---------------------------------------------------
# Each function returns 12 monthly amounts (None = blank cell).

def steady(price):
    return [price] * 12


def big_spender(price):
    # Business customers paying extra for add-ons some months.
    return [price + random.choice([0, 0, 150, 300, 800]) for _ in range(12)]


def slowly_dropping(price):
    # Spending shrinks a little most months (seats being removed, downgrades).
    amounts, current = [], price * random.uniform(1.3, 2.0)
    for _ in range(12):
        amounts.append(round(current, 2))
        current *= random.uniform(0.86, 0.97)
    return amounts


def stopped_paying(price):
    # Paid normally, then stopped. After that: zeros or blank cells.
    stop_at = random.randint(4, 10)
    return [price if i < stop_at else random.choice([0, None]) for i in range(12)]


def joined_late(price, start_index):
    # Blank before they signed up, then paying.
    return [None if i < start_index else price for i in range(12)]


# --- Build the 40 customers ---------------------------------------------

# How many customers of each type (adds up to 40).
PLAN_FOR_TYPE = {
    "big": ["Business"] * 5,
    "dropping": ["Pro", "Pro", "Business", "Pro", "Starter", "Pro"],
    "stopped": ["Starter", "Pro", "Starter", "Pro", "Starter"],
    "late": ["Starter", "Pro", "Starter", "Pro", "Business", "Starter", "Pro"],
    "steady": ["Starter"] * 9 + ["Pro"] * 8,
}

# A few customers get tricky values seen in the real export, so clean.py
# is tested on them. The numbers are customer numbers (2 means C-1002).
BIZ_SPELLINGS = {2: "Biz", 8: "biz ", 21: "BIZ"}     # short for Business
MISSPELLED_PLAN = {30: "Stater"}                      # not a real plan: should be flagged
DAYS_AGO_LOGINS = {12, 25, 33}                         # Last Login like "187 days ago"

customers = []
number = 1
for kind, plans in PLAN_FOR_TYPE.items():
    for plan in plans:
        first = random.choice(FIRST_NAMES)
        last = random.choice(LAST_NAMES)
        price = PLAN_PRICES[plan]
        if plan == "Business" and kind == "big":
            price = random.choice([1200, 1800, 2500])

        # Pick the spending pattern and a signup date that makes sense.
        if kind == "late":
            start = random.randint(2, 10)
            months = joined_late(price, start)
            signup = MONTH_STARTS[start] + timedelta(days=random.randint(0, 20))
        else:
            signup = date(2025, 9, 30) - timedelta(days=random.randint(30, 1100))
            months = {"big": big_spender, "dropping": slowly_dropping,
                      "stopped": stopped_paying, "steady": steady}[kind](price)

        # Customers who stopped paying haven't logged in for a while.
        if kind == "stopped":
            last_login = TODAY - timedelta(days=random.randint(60, 240))
            tickets = random.choice([0, 1, 4, 6])
        elif kind == "dropping":
            last_login = TODAY - timedelta(days=random.randint(10, 45))
            tickets = random.choice([2, 3, 5, 7])
        else:
            last_login = TODAY - timedelta(days=random.randint(0, 14))
            tickets = random.choice([0, 0, 1, 2])

        row = [
            f"C-{1000 + number}",
            messy_name(first, last),
            messy_email(first, last),
            messy_phone(number),
            messy_company(),
            pick_state(),
            random.choice(PLAN_SPELLINGS[plan]),
            messy_date(signup),
            *[messy_money(m) for m in months],
            messy_date(last_login),
            random.choice([tickets, str(tickets), f" {tickets} "]),
            random.choice(NOTES),
        ]
        # Swap in the tricky values (done after the random picks above, so
        # everything else in the file comes out the same as before).
        if number in BIZ_SPELLINGS:
            row[6] = BIZ_SPELLINGS[number]
        if number in MISSPELLED_PLAN:
            row[6] = MISSPELLED_PLAN[number]
        if number in DAYS_AGO_LOGINS:
            row[20] = f"{(TODAY - last_login).days} days ago"
        customers.append(row)
        number += 1

# Shuffle so the customer types aren't grouped together.
random.shuffle(customers)

# A few IDs get stray spaces or lowercase letters, like real exports.
customers[3][0] = f" {customers[3][0]}"
customers[11][0] = customers[11][0].lower()
customers[25][0] = f"{customers[25][0]} "

# The duplicate: same customer exported twice, with small differences.
duplicate = list(customers[7])
duplicate[2] = str(duplicate[2]).strip().upper()   # email in capitals
duplicate[22] = ""                                  # notes missing
customers.insert(20, duplicate)

# The blank row: an empty line in the middle of the data.
customers.insert(15, [None] * len(HEADERS))


# --- Save the file -------------------------------------------------------

wb = Workbook()
ws = wb.active
ws.title = "Customers"
ws.append(HEADERS)
for row in customers:
    ws.append(row)

# Show real date cells as dates (not as long numbers).
for row in ws.iter_rows(min_row=2):
    for cell in row:
        if isinstance(cell.value, datetime):
            cell.number_format = "m/d/yyyy"

wb.save(OUTPUT_FILE)
print(f"Saved {OUTPUT_FILE}: 40 fake customers, plus 1 duplicate and 1 blank row.")
