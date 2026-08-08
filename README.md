# holidayRecords

Compare currency **holiday dates** across multiple systems held in a single Excel
workbook, where **each tab = one system** (e.g. `WS`, `D3`, `Barracuda`).

You nominate one tab as the **source system**. The tool then builds the union of
every `(currency, date)` pair across all systems and reports **Y/N** for each
system per pair — surfacing matches and mismatches — and writes a **CSV report**.

## What the report looks like

| Currency | Date | WSS | Barracuda | Broadway | D3 | Status | Missing_In |
|----------|------|-----|-----------|----------|----|--------|------------|
| AED | 08/07/2026 | Y | Y | Y | Y | MATCH | |
| CNY | 05/01/2026 | Y | N | Y | Y | MISMATCH | Barracuda |
| JPY | 08/11/2026 | Y | Y | N | Y | MISMATCH | Broadway |
| KRW | 08/17/2026 | N | Y | N | N | MISMATCH | WSS; Broadway; D3 |

- The **source column** comes first, then the other systems.
- `Missing_In` lists every system that does **not** have that currency/date.
- A row where the source is `N` means another system has a date the source lacks.

## Expected Excel format

Each **tab is one system**, and the tab/sheet name **is** the system name — so a
`System` column inside the sheet is ignored (you don't need to remove it). Each
tab lists currencies (3-letter codes: `CNY`, `USD`, `HKD` …) and holiday dates.

**Standard layout** — `System` (optional, ignored), `Holiday Ccy`, `Holiday Date`:

| System | Holiday Ccy | Holiday Date |
|--------|-------------|--------------|
| WSS | CNY | 01/01/2026 |
| WSS | CNY | 02/16/2026 |
| WSS | PEN | 08/06/2026 |

The currency column is recognised from headers like `Holiday Ccy`, `Currency`,
`Ccy`; the date column from `Holiday Date`, `Date`. A `WIDE` layout (currency
codes as column headers, dates listed beneath) is also auto-detected. Real Excel
date cells are accepted as well as text.

### Date format

Dates default to **MM/DD/YYYY**. To use a different format, pass `--date-format`
locally or set the **Date format** input when running the workflow — e.g.
`DD/MM/YYYY` or `YYYY-MM-DD`. The chosen format is used both to interpret text
dates in the workbook and to write dates in the report.

## Run it via GitHub Actions (no local setup)

1. Put your workbook in the `input/` folder (commit it), e.g. `input/holidays.xlsx`.
2. Go to the **Actions** tab → **Holiday Comparison Report** → **Run workflow**.
3. Enter the **source system** (tab name, e.g. `WSS`). Optionally set a specific
   file and/or the date format.
4. The report is uploaded as a downloadable **artifact** and committed to
   `output/holiday_comparison.csv`.

## Run it locally

```bash
pip install -r requirements.txt

# generate a demo workbook (optional)
python scripts/make_sample.py

# compare — uses the newest file in input/ if --file is omitted
python compare_holidays.py --source WSS
python compare_holidays.py --file input/holidays.xlsx --source WSS --output output/holiday_comparison.csv

# use a different date format (default is MM/DD/YYYY)
python compare_holidays.py --source WSS --date-format DD/MM/YYYY
```

The source name is matched case-insensitively; if it isn't found, the available
tab names are listed.
