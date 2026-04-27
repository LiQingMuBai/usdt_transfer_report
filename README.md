# print_transfer_report

A small Rust CLI that reads a blockchain transfer export CSV (`Transfers.csv`), converts timestamps from **UTC to Beijing Time (UTC+8)**, aggregates transfers **by day (BJT)**, and exports an **Excel (.xlsx)** report with a **dark-themed column chart**.

## Input CSV

The tool expects a CSV file with headers like:

- `Txn Hash`
- `Block`
- `Time(UTC)`
- `From`
- `To`
- `Token`
- `Token Symbol`
- `Amount/TokenID`
- `Result`
- `Status`

## Quick Start

```bash
cd /Users/masion/monitor/print_transfer_report
cargo run -- --summary-only
```

This reads configuration from `.env` (see below), prints the daily summary to stdout, and writes the Excel report to the configured output path (via `XLSX_PATH`).

## Configuration (.env)

Create a `.env` file in the project root (you can copy from `.env.example`).

Required:

- `CSV_PATH`: absolute path to `Transfers.csv`
- `XLSX_PATH`: output `.xlsx` file path (including filename)

Range:

- `RANGE_DAYS`: number of days to include (integer > 0). If not set, all days are included.

Defaults / behavior:

- The report **excludes “today” (Beijing date)** by default, to avoid partial-day data.

Example:

```ini
CSV_PATH=/Users/masion/Desktop/Transfers.csv
XLSX_PATH=/Users/masion/monitor/print_transfer_report/transfer_daily_summary.xlsx
RANGE_DAYS=40
```

## CLI Options

- `--summary-only`: do not print each transfer record; only print the daily summary.
- `--xlsx [PATH]`:
  - if `PATH` is provided, write the Excel report to that path
  - if `PATH` is omitted, use `XLSX_PATH` from `.env` (or a local default filename)
- `--no-xlsx`: disable Excel output for this run.
- `--days N`: override `RANGE_DAYS` from `.env` for this run.
- `--include-today`: include today (Beijing date) in the output (default is excluded).

Examples:

```bash
# Use .env (recommended)
cargo run -- --summary-only

# Override the day range for a single run
cargo run -- --summary-only --days 60

# Write to a custom path (overrides .env)
cargo run -- --summary-only --xlsx /tmp/transfer_daily_summary.xlsx

# Include today if you want the partial-day data
cargo run -- --summary-only --include-today

# No Excel output
cargo run -- --summary-only --no-xlsx
```

## Output

The generated Excel file contains:

- Worksheet: `汇总`
- Columns:
  - `日期(BJT)` (Beijing date)
  - `转账次数` (transfer count)
- Column chart:
  - Type: Column
  - Color: dark blue (`#1F4E79`)
  - Title: `每日转账次数（北京时间）`

## Notes

- Timestamp parsing supports:
  - `YYYY-MM-DD HH:MM:SS`
  - `YYYY-MM-DD HH:MM`
- If you pipe output to tools like `head`, the program exits cleanly (handles broken pipe).
