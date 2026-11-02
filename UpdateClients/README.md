# Client Pivot & Plan Tracker

A Python script that automates daily reporting on a logistics company's client base. It merges raw Excel exports, calculates each client's current status, builds a monthly pivot table, distributes a monthly sales plan across clients, forecasts plan completion, and publishes the results to Google Sheets.

## Features

- **Merge & clean**: combines all `.xlsx` exports from a folder, removes duplicate waybills, parses dates, and skips the summary row in each file.
- **Current client status**: determines each client's latest status from shipments over the last 90 days.
- **Attribute normalization**: takes client attributes (ABC class, parent counterparty, business segment, sales and support managers) from the most recent shipment, so each client has one consistent record.
- **Monthly pivot**: revenue, weight, number of packages, and unique waybills per client per month.
- **Plan distribution**: spreads a fixed monthly plan across clients in proportion to their revenue over the previous 3 months. Clients with the status "new" and negative bases are excluded.
- **Weighted forecast**: extrapolates month-to-date results using working-day weights (Mon-Fri = 1, Sat/Sun = 0.25) and shows the change against the previous month.
- **Manager scorecards**: fact, plan, forecast, and forecast-to-plan % for sales managers and support managers, plus a total row. Cells are green at or above 90% and red below it.
- **Excel and Google Sheets output**: formatted workbooks with merged headers, color-coded month blocks, frozen panes, and filters. The Google Sheets upload applies the same formatting through the Sheets API, including a red/green gradient for month-over-month change.

## Tech stack

Python 3.10+, pandas, openpyxl, gspread, google-auth

## Installation

```bash
pip install pandas openpyxl gspread google-auth
```

## Configuration

Edit the constants at the top of `update_clients.py`:

| Setting | Description |
|---|---|
| `INPUT_DIR` | Folder with the source `.xlsx` exports |
| `OUTPUT`, `PIVOT_OUTPUT`, `BASE_OUTPUT` | Paths of the generated Excel files |
| `GOOGLE_CREDENTIALS` | Path to the Google service account JSON key |
| `GOOGLE_SHEET_URL` | URL of the target Google spreadsheet |
| `PLAN_AMOUNT` | Monthly plan in UAH |
| `PLAN_MONTH` | Plan month (1-12); `None` means the current month |
| `PLAN_BASE_MONTHS` | Number of previous months used as the distribution base |
| `PLAN_EXCLUDE_STATUSES` | Client statuses that receive no plan |
| `TOTAL_MANAGER` | Manager under whose name the overall total row is shown |

Google Sheets setup: create a service account, download its JSON key, and share the target spreadsheet with the service account's email (Editor access). **Never commit the credentials file to the repository.**

## Usage

```bash
python update_clients.py
python update_clients.py --input-dir "path/to/reports"
python update_clients.py --plan 10000000 --plan-month 11
python update_clients.py --no-google
```

| Argument | Description |
|---|---|
| `--input-dir` | Folder with source files (overrides `INPUT_DIR`) |
| `--plan` | Monthly plan amount; if omitted, the default from the config is used |
| `--plan-month` | Month number for the plan and forecast |
| `--no-google` | Build Excel files only, skip the Google Sheets upload |

## Input format

Each source file is expected to have 4 header rows above the table and a summary row at the end. Required columns include: waybill number, shipment date (`dd.mm.yyyy`), payer counterparty (code, name, EDRPOU), client status, ABC class, business segment, sales manager, support manager, cost with VAT, weight, and number of packages. Column names are in Ukrainian.

## Output

**Excel files**
- `Active_clients2026(base).xlsx`: monthly pivot with forecast
- `Active_clients2026(pivot).xlsx`: monthly pivot with plan and forecast

**Google Sheets worksheets**
- `Поточна динаміка`: pivot with forecast (no plan)
- `План(поточний місяць)`: pivot with plan and forecast
- `Виконання(поточний місяць)`: manager scorecards for sales and support, plus the overall total

A log file `update_clients.log` is written next to the output files.

## Scheduling

The script is designed to run daily. On Windows, use Task Scheduler to run `python update_clients.py` each morning, after the previous day's data has been exported.
