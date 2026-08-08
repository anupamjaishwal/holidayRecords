# holidayRecords

Compare currency **holiday dates** across multiple systems held in a single Excel
workbook, where **each tab = one system** (e.g. `WS`, `D3`, `Barracuda`).

You nominate one tab as the **source system**. The tool then builds the union of
every `(currency, date)` pair across all systems and reports **Y/N** for each
system per pair — surfacing matches and mismatches — and writes a **CSV report**.

## What the report looks like

| Currency | Date | WS | D3 | Barracuda | Status | Missing_In |
|----------|------|----|----|-----------|--------|------------|
| CAD | 01/01/2020 | Y | Y | Y | MATCH | |
| CAD | 12/25/2020 | Y | N | Y | MISMATCH | D3 |
| USD | 07/04/2020 | Y | Y | N | MISMATCH | Barracuda |
| USD | 11/26/2020 | N | Y | N | MISMATCH | WS; Barracuda |

- The **source column** comes first, then the other systems.
- `Missing_In` lists every system that does **not** have that currency/date.
- A row where the source is `N` means another system has a date the source lacks.

## Expected Excel format

Each tab lists currencies (3-letter codes: `CAD`, `USD`, `HKD` …) and holiday
dates in **MM/DD/YYYY**. Two layouts are auto-detected:

**LONG** (recommended) — a `Currency` column and a `Holiday Date` column:

| Currency | Holiday Date |
|----------|--------------|
| CAD | 01/01/2020 |
| CAD | 07/01/2020 |
| USD | 01/01/2020 |

**WIDE** — currency codes as headers, dates down each column:

| CAD | USD | HKD |
|-----|-----|-----|
| 01/01/2020 | 01/01/2020 | 01/01/2020 |
| 07/01/2020 | | 10/01/2020 |

Real Excel date cells are also accepted (not only text).

## Run it via GitHub Actions (no local setup)

1. Put your workbook in the `input/` folder (commit it), e.g. `input/holidays.xlsx`.
2. Go to the **Actions** tab → **Holiday Comparison Report** → **Run workflow**.
3. Enter the **source system** (tab name, e.g. `WS`). Optionally set a specific file.
4. The report is uploaded as a downloadable **artifact** and committed to
   `output/holiday_comparison.csv`.

## Run it locally

```bash
pip install -r requirements.txt

# generate a demo workbook (optional)
python scripts/make_sample.py

# compare — uses the newest file in input/ if --file is omitted
python compare_holidays.py --source WS
python compare_holidays.py --file input/holidays.xlsx --source WS --output output/holiday_comparison.csv
```

The source name is matched case-insensitively; if it isn't found, the available
tab names are listed.
