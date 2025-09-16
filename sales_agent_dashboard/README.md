# Sales Agent Excel Dashboard

Generate an Excel dashboard for sales agents from a CSV dataset using Python and pandas + XlsxWriter.

## Quick start

1) Create a virtual environment (optional) and install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2) Run the generator with the provided sample data:

```bash
python generate_dashboard.py --input data/sample_sales.csv --output sales_dashboard.xlsx
```

3) Open `sales_dashboard.xlsx` in Excel. You'll find:

- Dashboard: KPI cards and charts (Monthly Revenue, Top Agents, Category Share, Region Revenue)
- Aggregates: Pivot-like summary tables that power charts
- RawData: Original data with filters and frozen header

## Input CSV format

Required columns:

- OrderDate (YYYY-MM-DD)
- Agent
- Region
- Product
- Category
- Units (int)
- UnitPrice (float)
- Revenue (float) — if omitted, it's computed as Units * UnitPrice

## Customize

- To control Top N agents chart: `--top 5`
- Use your own data: `--input /path/to/your.csv`
- Change the output path: `--output /path/to/your.xlsx`

## Notes

- Slicers and native Excel PivotTables are not created; charts are based on pre-aggregated tables for portability.
- Open the file in desktop Excel for the best chart rendering.

