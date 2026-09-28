# monday_summary.py
# Writes reports/monday-summary.txt: a plain-English summary for sharing.
# It shows totals only, never individual customers, and hides any group
# with fewer than 3 customers.
#
# What's in it: revenue for each of the last 12 months, the change from
# the month before, new customers this month, customers who stopped
# paying, and revenue by plan.
#
# Run clean.py first (run_all.py does both). Then:
#   python3 monday_summary.py path/to/customers.xlsx

import sys
from datetime import date

from report_tools import (SMALLEST_GROUP, dollars, find_files, is_paying,
                          load_customers, month_name, protect_groups)


def how_many(count):
    """Customer counts in words. Counts of 1 or 2 are shown as
    "fewer than 3", so a single customer can't be singled out."""
    if count == 0:
        return "none"
    if count < SMALLEST_GROUP:
        return f"fewer than {SMALLEST_GROUP}"
    return str(count)


def change_text(now, before):
    """ "up $1,200 (4.5%)" or "down $300 (-1.2%)" or "no change"."""
    # Round first, so the change matches the whole-dollar figures shown.
    now, before = round(now), round(before)
    difference = now - before
    if difference == 0:
        return "no change"
    direction = "up" if difference > 0 else "down"
    percent = f" ({difference / before:+.1%})" if before else ""
    return f"{direction} {dollars(abs(difference))}{percent}"


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python3 monday_summary.py path/to/customers.xlsx")
    clean_file, reports_folder = find_files(sys.argv[1])
    months, customers = load_customers(clean_file)
    latest, previous = len(months) - 1, len(months) - 2
    this_month = months[latest]

    # Revenue and number of paying customers for each month.
    revenue = [sum(c["amounts"][i] or 0 for c in customers if is_paying(c, i))
               for i in range(len(months))]
    payers = [sum(1 for c in customers if is_paying(c, i)) for i in range(len(months))]
    shown = [n >= SMALLEST_GROUP for n in payers]   # privacy: hide tiny months

    # New this month: signed up during the latest month. (If a signup date
    # is missing, their first month on the books is used instead.)
    def joined_this_month(c):
        if c["signup"]:
            return (c["signup"].year, c["signup"].month) == (this_month.year, this_month.month)
        first = next((i for i, a in enumerate(c["amounts"]) if a is not None), None)
        return first == latest
    new_count = sum(1 for c in customers if joined_this_month(c))

    # Stopped paying: paid last month, paid nothing this month.
    stopped_count = sum(1 for c in customers
                        if previous >= 0 and is_paying(c, previous) and not is_paying(c, latest))
    # Also: everyone who paid at some point this year but not this month.
    lapsed_count = sum(1 for c in customers
                       if any(is_paying(c, i) for i in range(latest)) and not is_paying(c, latest))

    # Revenue by plan, this month and over all 12 months.
    def plan_table(amount_for):
        groups = {}
        for c in customers:
            amount = amount_for(c)
            if amount > 0:
                groups.setdefault(c["plan"], []).append(amount)
        table = [(label, sum(items), len(items))
                 for label, items in protect_groups(groups, "Other plans")]
        return sorted(table, key=lambda row: -row[1])

    plans_now = plan_table(lambda c: c["amounts"][latest] or 0)
    plans_year = plan_table(lambda c: sum(a for a in c["amounts"] if a))
    year_total = sum(r for r, ok in zip(revenue, shown) if ok)

    # --- Write it out in plain English -------------------------------------
    lines = []
    add = lines.append
    add("MONDAY CUSTOMER SUMMARY")
    add(f"Prepared {date.today():%B %d, %Y}. Covers {month_name(months[0])} "
        f"to {month_name(this_month)}.")
    add("")
    add("HEADLINES")
    if shown[latest]:
        headline = f"- {month_name(this_month)} revenue was {dollars(revenue[latest])}"
        if previous >= 0 and shown[previous]:
            headline += (f", {change_text(revenue[latest], revenue[previous])} "
                         f"from {month_name(months[previous])}")
        add(headline + ".")
        add(f"- {how_many(payers[latest]).capitalize()} customers paid this month.")
    else:
        add(f"- {month_name(this_month)} revenue isn't shown: fewer than "
            f"{SMALLEST_GROUP} customers paid.")
    add(f"- New customers this month: {how_many(new_count)}.")
    add(f"- Customers who stopped paying this month: {how_many(stopped_count)}.")
    add(f"- Customers who paid earlier in the year but not this month: {how_many(lapsed_count)}.")
    add(f"- Total revenue over the last {len(months)} months: {dollars(year_total)}.")
    add("")

    add("REVENUE BY MONTH")
    add(f"  {'Month':<16}{'Revenue':>12}   Change from the month before")
    for i, month in enumerate(months):
        if not shown[i]:
            add(f"  {month_name(month):<16}{'(hidden)':>12}   fewer than {SMALLEST_GROUP} paying customers")
            continue
        change = "-"
        if i > 0 and shown[i - 1]:
            change = change_text(revenue[i], revenue[i - 1])
        add(f"  {month_name(month):<16}{dollars(revenue[i]):>12}   {change}")
    add("")

    add("REVENUE BY PLAN")
    width = max([len(row[0]) for row in plans_now + plans_year] + [10]) + 2
    for title, table in [(month_name(this_month), plans_now),
                         (f"Last {len(months)} months", plans_year)]:
        total = sum(row[1] for row in table)
        add(f"  {title}:")
        if not table:
            add("    (not shown: too few customers)")
        for label, amount, count in table:
            share = f"{amount / total:.0%}" if total else "-"
            add(f"    {label:<{width}}{dollars(amount):>12}   {share:>4} of revenue, "
                f"{count} customers")
    add("")

    add("HOW TO READ THIS")
    add(f"- \"New\" means the customer signed up during {month_name(this_month)}.")
    add(f"- \"Stopped paying\" means they paid in {month_name(months[previous])} "
        f"but nothing in {month_name(this_month)}.")
    add(f"- To protect customer privacy, no individual customers are listed, and any "
        f"group with fewer than {SMALLEST_GROUP} customers is combined with others "
        "or left out.")

    out_file = reports_folder / "monday-summary.txt"
    out_file.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Saved {out_file}")


if __name__ == "__main__":
    main()
