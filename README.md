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

**Optional description columns:** if the **source tab** has `EventDayOfWeek`
and/or `EventName` columns, their values are carried into the report (right
after the `Date` column). These are read from the source tab only — if other
tabs don't have them, or the source tab doesn't either, nothing fails; the
columns simply don't appear.

### Date format

Dates default to **MM/DD/YYYY**. To use a different format, pass `--date-format`
locally or set the **Date format** input when running the workflow — e.g.
`DD/MM/YYYY` or `YYYY-MM-DD`. The chosen format is used both to interpret text
dates in the workbook and to write dates in the report.

## Use it in the browser — no login, no install (recommended)

A self-contained web tool lives in `docs/`. It reads the Excel file **entirely in
your browser** — nothing is uploaded to any server, and no GitHub account is
needed to use it. This is the option for a locked-down / work machine.

**Two ways to open it:**

- **Hosted (public URL):** once GitHub Pages is enabled (see below), anyone can
  visit `https://anupamjaishwal.github.io/holidayRecords/` and use it — no login.
- **Offline:** download the repo (green **Code → Download ZIP**), unzip, and open
  `docs/index.html` in a browser. Works with no network at all.

**Steps:** pick your `.xlsx` → choose the **source system** → (optional) set the
date format → (optional) tick the **currencies to compare** → **Compare** →
**Download CSV**. Toggle *Show mismatches only* to focus on differences.

After a file is loaded, every currency found in the workbook appears as a
checkbox (all selected by default). Untick any you don't care about — the report
then shows only the currencies you keep selected. Use **Select all** / **Clear**
to toggle them in bulk.

> Requires a reasonably modern browser (Chrome/Edge/Firefox/Safari from ~2023+),
> which is used to unzip the `.xlsx` locally.

### One-time: enable the public URL (owner only)

The repo owner does this once (it does require logging in, but only for setup —
afterwards visitors need no login):

1. **Settings → Pages**
2. **Build and deployment → Source:** *Deploy from a branch*
3. **Branch:** `main`, **Folder:** `/docs` → **Save**
4. Wait ~1 minute; the site appears at `https://anupamjaishwal.github.io/holidayRecords/`.

## Run it via GitHub Actions (owner only — requires login)

1. Put your workbook in the `input/` folder (commit it) — **any filename works**,
   e.g. `input/calendar_2026.xlsx`.
2. Go to the **Actions** tab → **Holiday Comparison Report** → **Run workflow**.
3. Enter the **source system** (tab name, e.g. `WSS`). In the file field, type the
   filename (e.g. `calendar_2026.xlsx`) — or leave it blank to use the newest file
   in `input/`. Optionally set the date format, and optionally list the
   **currencies** to include (e.g. `USD,CAD,HKD`; blank = all).
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

# only compare specific currencies (omit --currencies to include all)
python compare_holidays.py --source WSS --currencies USD,CAD,HKD
```

The source name is matched case-insensitively; if it isn't found, the available
tab names are listed.
