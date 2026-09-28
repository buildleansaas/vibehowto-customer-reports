# top_customers.py
# Ranks customers by how much they spent over the last 12 months and saves
# reports/top-customers.csv. Customers are shown by Customer ID only.
#
# For each customer: Customer ID, plan, state, average monthly spend
# (counting only the months they were a customer), and 12-month total.
# At the bottom: average spend for each plan.
#
# Run clean.py first (run_all.py does both). Then:
#   python3 top_customers.py path/to/customers.xlsx

import csv
import sys

from report_tools import (SMALLEST_GROUP, customer_months, find_files,
                          load_customers, month_name, protect_groups)


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python3 top_customers.py path/to/customers.xlsx")
    clean_file, reports_folder = find_files(sys.argv[1])
    months, customers = load_customers(clean_file)

    # Work out each customer's total and monthly average.
    ranked = []
    for c in customers:
        total = sum(a for a in c["amounts"] if a)
        active = customer_months(c)
        average = sum(c["amounts"][i] or 0 for i in active) / len(active) if active else None
        ranked.append({**c, "total": total, "average": average, "months": len(active)})

    # Highest 12-month total first (ties sorted by Customer ID).
    ranked.sort(key=lambda r: (-r["total"], str(r["id"])))

    out_file = reports_folder / "top-customers.csv"
    period = f"{month_name(months[0])} to {month_name(months[-1])}"
    with open(out_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        # Part 1: one line per customer.
        writer.writerow(["Rank", "Customer ID", "Plan", "State",
                         "Average monthly spend", f"Total spend ({period})",
                         "Months counted"])
        for rank, r in enumerate(ranked, start=1):
            writer.writerow([rank, r["id"], r["plan"], r["state"],
                             "" if r["average"] is None else f"{r['average']:.2f}",
                             f"{r['total']:.2f}", r["months"]])

        # Part 2: average spend by plan. Only customers who paid at least
        # once are counted. Plans with fewer than 3 customers are combined.
        by_plan = {}
        for r in ranked:
            if r["months"]:
                by_plan.setdefault(r["plan"], []).append(r)
        groups = protect_groups(by_plan, "Other plans")
        groups.sort(key=lambda g: -sum(r["average"] for r in g[1]) / len(g[1]))
        everyone = [r for r in ranked if r["months"]]

        writer.writerow([])
        writer.writerow(["Average spend by plan"])
        writer.writerow(["Plan", "Customers", "Average monthly spend",
                         "Average 12-month total"])
        for label, members in groups + [("All customers", everyone)]:
            if len(members) < SMALLEST_GROUP:
                continue
            writer.writerow([label, len(members),
                             f"{sum(r['average'] for r in members) / len(members):.2f}",
                             f"{sum(r['total'] for r in members) / len(members):.2f}"])
        writer.writerow([])
        writer.writerow([f"Plans with fewer than {SMALLEST_GROUP} customers are "
                         "combined into 'Other plans' for privacy."])

    print(f"Saved {out_file} ({len(ranked)} customers)")


if __name__ == "__main__":
    main()
