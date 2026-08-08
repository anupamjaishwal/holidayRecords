#!/usr/bin/env python3
"""Generate a sample holidays.xlsx (LONG layout) for testing/demo.

Creates one tab per system. WS is intended to be used as the source.
Run:  python scripts/make_sample.py
"""
import os
from openpyxl import Workbook

# system -> list of (currency, "MM/DD/YYYY")
DATA = {
    "WS": [  # source system
        ("CAD", "01/01/2020"), ("CAD", "07/01/2020"), ("CAD", "12/25/2020"),
        ("USD", "01/01/2020"), ("USD", "07/04/2020"),
        ("HKD", "01/01/2020"), ("HKD", "10/01/2020"),
    ],
    "D3": [
        ("CAD", "01/01/2020"), ("CAD", "07/01/2020"),           # missing 12/25
        ("USD", "01/01/2020"), ("USD", "07/04/2020"), ("USD", "11/26/2020"),  # extra 11/26
        ("HKD", "01/01/2020"), ("HKD", "10/01/2020"),
    ],
    "Barracuda": [
        ("CAD", "01/01/2020"), ("CAD", "07/01/2020"), ("CAD", "12/25/2020"),
        ("USD", "01/01/2020"),                                   # missing 07/04
        ("HKD", "01/01/2020"),                                   # missing 10/01
    ],
}


def main():
    wb = Workbook()
    wb.remove(wb.active)
    for system, rows in DATA.items():
        ws = wb.create_sheet(title=system)
        ws.append(["Currency", "Holiday Date"])
        for currency, date in rows:
            ws.append([currency, date])
    os.makedirs("input", exist_ok=True)
    out = "input/holidays.xlsx"
    wb.save(out)
    print(f"Wrote {out} with tabs: {', '.join(DATA)}")


if __name__ == "__main__":
    main()
