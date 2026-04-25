# print_report

Generate a daily summary Excel report from `resource-orders.csv`.

The program:
- Groups data by **Dispatch Time** (derived from `Delegation Time (UTC+8)` date).
- Excludes **today** (UTC+8) from the summary.
- Produces an `.xlsx` with:
  - A summary table (per day + a total row).
  - A column chart (dark color palette).
- Shows a terminal progress UI while running.

## Input

Expected CSV headers (minimum required):
- `Delegation Time (UTC+8)`
- `Resource Quantity`

## Mapping Rules

`Resource Quantity` is mapped into two types:
- `65000` → **With-U Energy Trades** (有U能量笔数) = 1 trade
- `131000` → **No-U Energy Trades** (无U能量笔数) = 1 trade

In the Excel output:
- **无U能量笔数** is displayed as **(No-U trades × 2)** as requested.
- **每日转账次数** (Daily transfer count) = With-U trades + No-U trades (trade count basis).

## Configuration (.env)

Create a `.env` file in this folder (or copy from `.env.example`):

```bash
cp .env.example .env
```

Available keys:

```env
INPUT_CSV=/Users/masion/Desktop/resource-orders.csv
OUTPUT_XLSX=delegation_summary.xlsx
```

Priority order:
1. CLI arguments
2. `.env` values
3. Built-in defaults

## Run

From the project directory:

```bash
cargo run --bin delegation_summary_xlsx
```

Override input/output via CLI:

```bash
cargo run --bin delegation_summary_xlsx -- /path/to/resource-orders.csv /path/to/output.xlsx
```

## Output

The generated workbook contains a sheet named `汇总` with columns:
- 派发时间
- 有U能量笔数
- 无U能量笔数
- 每日转账次数

And a last row `总计` with SUM formulas for each numeric column.

## Notes (macOS)

If you set `OUTPUT_XLSX` to write into protected folders (e.g. Desktop/Documents) and get `PermissionDenied`, grant your terminal app permission in:
System Settings → Privacy & Security → Files and Folders.

