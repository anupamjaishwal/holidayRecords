#!/usr/bin/env python3
"""Generate a sample holidays.xlsx matching the real layout, for testing/demo.

Real layout (per tab): columns  System | Holiday Ccy | Holiday Date
The "System" column repeats the tab name and is IGNORED by the tool — the
system name comes from the tab (sheet) name. WSS is intended as the source.

Run:  python scripts/make_sample.py
"""
import os
from openpyxl import Workbook

# tab (system) -> list of (currency, "MM/DD/YYYY")
DATA = {
    "WSS": [  # source system
        ("CNY", "01/01/2026"), ("CNY", "02/16/2026"), ("CNY", "05/01/2026"),
        ("PEN", "08/06/2026"),
        ("AED", "08/07/2026"),
        ("JPY", "08/11/2026"),
        ("INR", "08/17/2026"),
    ],
    "Barracuda": [
        ("CNY", "01/01/2026"), ("CNY", "02/16/2026"),          # missing 05/01
        ("PEN", "08/06/2026"),
        ("AED", "08/07/2026"),
        ("JPY", "08/11/2026"),
        ("INR", "08/17/2026"), ("KRW", "08/17/2026"),          # extra KRW
    ],
    "Broadway": [
        ("CNY", "01/01/2026"), ("CNY", "02/16/2026"), ("CNY", "05/01/2026"),
        ("AED", "08/07/2026"),                                  # missing PEN
        ("INR", "08/17/2026"),                                  # missing JPY
    ],
    "D3": [
        ("CNY", "01/01/2026"), ("CNY", "05/01/2026"),          # missing 02/16
        ("PEN", "08/06/2026"),
        ("AED", "08/07/2026"),
        ("JPY", "08/11/2026"),
        ("INR", "08/17/2026"),
    ],
}


def main():
    wb = Workbook()
    wb.remove(wb.active)
    for system, rows in DATA.items():
        ws = wb.create_sheet(title=system)
        ws.append(["System", "Holiday Ccy", "Holiday Date"])
        for currency, date in rows:
            ws.append([system, currency, date])
    os.makedirs("input", exist_ok=True)
    out = "input/holidays.xlsx"
    wb.save(out)
    print(f"Wrote {out} with tabs: {', '.join(DATA)}")


if __name__ == "__main__":
    main()
