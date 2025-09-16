#!/usr/bin/env python3
"""
Generate an Excel sales dashboard for sales agents from a CSV file.

Input CSV expected columns (case-sensitive):
  - OrderDate (YYYY-MM-DD or ISO datetime)
  - Agent
  - Region
  - Product
  - Category
  - Units (int)
  - UnitPrice (float)
  - Revenue (float)  # If missing, will be computed as Units * UnitPrice

Output: Excel file with sheets:
  - RawData: all source rows
  - Aggregates: pivot-like summary tables for charts
  - Dashboard: KPIs and charts
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Tuple

import pandas as pd


DEFAULT_OUTPUT_FILENAME = "sales_dashboard.xlsx"


@dataclass
class DashboardData:
    raw: pd.DataFrame
    monthly: pd.DataFrame
    by_agent: pd.DataFrame
    top_agents: pd.DataFrame
    by_category: pd.DataFrame
    by_region: pd.DataFrame
    kpis: dict


def read_sales_csv(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(csv_path)
    if "OrderDate" not in df.columns:
        raise ValueError("Input CSV must include an 'OrderDate' column")
    # Parse dates
    df["OrderDate"] = pd.to_datetime(df["OrderDate"], errors="coerce")
    if df["OrderDate"].isna().any():
        raise ValueError("Some 'OrderDate' values could not be parsed as dates")

    # Standardize text columns if present
    for col in ["Agent", "Region", "Product", "Category"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Ensure numeric columns
    if "Units" in df.columns:
        df["Units"] = pd.to_numeric(df["Units"], errors="coerce").fillna(0).astype(int)
    if "UnitPrice" in df.columns:
        df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="coerce").fillna(0.0)
    if "Revenue" not in df.columns:
        if {"Units", "UnitPrice"}.issubset(df.columns):
            df["Revenue"] = df["Units"] * df["UnitPrice"]
        else:
            raise ValueError("CSV missing 'Revenue' and cannot compute without 'Units' and 'UnitPrice'")

    # Sort by date for readability
    df = df.sort_values("OrderDate").reset_index(drop=True)
    return df


def build_aggregates(df: pd.DataFrame, top_n_agents: int = 10) -> DashboardData:
    # Monthly totals
    monthly = (
        df.set_index("OrderDate")
        .groupby(pd.Grouper(freq="MS"))["Revenue"]
        .sum()
        .reset_index()
    )
    monthly["Month"] = monthly["OrderDate"].dt.strftime("%Y-%m")
    monthly = monthly[["Month", "Revenue"]]

    # By agent totals
    by_agent = df.groupby("Agent", dropna=False)["Revenue"].sum().reset_index()
    by_agent = by_agent.sort_values("Revenue", ascending=False).reset_index(drop=True)
    top_agents = by_agent.head(top_n_agents).reset_index(drop=True)

    # Category totals
    if "Category" in df.columns:
        by_category = df.groupby("Category", dropna=False)["Revenue"].sum().reset_index()
    else:
        by_category = pd.DataFrame({"Category": ["All"], "Revenue": [df["Revenue"].sum()]})
    by_category = by_category.sort_values("Revenue", ascending=False).reset_index(drop=True)

    # Region totals
    if "Region" in df.columns:
        by_region = df.groupby("Region", dropna=False)["Revenue"].sum().reset_index()
    else:
        by_region = pd.DataFrame({"Region": ["All"], "Revenue": [df["Revenue"].sum()]})
    by_region = by_region.sort_values("Revenue", ascending=False).reset_index(drop=True)

    # KPIs
    total_revenue = float(df["Revenue"].sum())
    num_agents = int(df["Agent"].nunique()) if "Agent" in df.columns else 0
    total_orders = int(len(df))
    avg_deal_size = float(df["Revenue"].mean()) if total_orders > 0 else 0.0

    kpis = {
        "Total Revenue": total_revenue,
        "Total Orders": total_orders,
        "# of Agents": num_agents,
        "Avg Deal Size": avg_deal_size,
    }

    return DashboardData(
        raw=df,
        monthly=monthly,
        by_agent=by_agent,
        top_agents=top_agents,
        by_category=by_category,
        by_region=by_region,
        kpis=kpis,
    )


def write_excel_dashboard(data: DashboardData, output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        # Write data sheets
        data.raw.to_excel(writer, sheet_name="RawData", index=False)

        # Aggregates sheet layout positions
        sheet_agg = "Aggregates"
        start_positions = {
            "Monthly": (0, 0),
            "ByAgent": (0, 4),
            "ByCategory": (0, 8),
            "ByRegion": (0, 12),
        }

        # Prepare aggregate DataFrames with column headers
        monthly_df = data.monthly.rename(columns={"Revenue": "Revenue"})
        by_agent_df = data.by_agent.rename(columns={"Revenue": "Revenue"})
        top_agents_df = data.top_agents.rename(columns={"Revenue": "Revenue"})
        by_category_df = data.by_category.rename(columns={"Revenue": "Revenue"})
        by_region_df = data.by_region.rename(columns={"Revenue": "Revenue"})

        # Write aggregates
        monthly_df.to_excel(
            writer, sheet_name=sheet_agg, startrow=start_positions["Monthly"][0], startcol=start_positions["Monthly"][1], index=False
        )
        by_agent_df.to_excel(
            writer, sheet_name=sheet_agg, startrow=start_positions["ByAgent"][0], startcol=start_positions["ByAgent"][1], index=False
        )
        by_category_df.to_excel(
            writer, sheet_name=sheet_agg, startrow=start_positions["ByCategory"][0], startcol=start_positions["ByCategory"][1], index=False
        )
        by_region_df.to_excel(
            writer, sheet_name=sheet_agg, startrow=start_positions["ByRegion"][0], startcol=start_positions["ByRegion"][1], index=False
        )

        workbook = writer.book
        ws_dashboard = workbook.add_worksheet("Dashboard")

        # Formats
        fmt_title = workbook.add_format({"bold": True, "font_size": 20})
        fmt_kpi_label = workbook.add_format({"bold": True, "font_size": 11, "align": "center"})
        fmt_kpi_value_currency = workbook.add_format({"bold": True, "font_size": 16, "align": "center", "num_format": "$#,##0"})
        fmt_kpi_value_int = workbook.add_format({"bold": True, "font_size": 16, "align": "center", "num_format": "#,##0"})
        fmt_box = workbook.add_format({"bg_color": "#F2F2F2", "border": 1})

        # Title
        ws_dashboard.merge_range("A1:H1", "Sales Agent Dashboard", fmt_title)

        # KPI cards grid 2x2
        kpi_items = list(data.kpis.items())
        # Ensure order: Total Revenue, Total Orders, # of Agents, Avg Deal Size
        key_order = ["Total Revenue", "Total Orders", "# of Agents", "Avg Deal Size"]
        ordered_kpis = [(k, data.kpis[k]) for k in key_order]

        kpi_positions = [(2, 0), (2, 4), (7, 0), (7, 4)]  # (row, col) top-left of each card
        for (label, value), (r, c) in zip(ordered_kpis, kpi_positions):
            ws_dashboard.write_blank(r, c, None, fmt_box)
            ws_dashboard.write_blank(r, c + 3, None, fmt_box)
            # Draw box by formatting a 4x4 area
            for rr in range(r, r + 4):
                for cc in range(c, c + 4):
                    ws_dashboard.write_blank(rr, cc, None, fmt_box)
            ws_dashboard.merge_range(r, c, r, c + 3, label, fmt_kpi_label)
            # Value row
            if label in ("Total Revenue", "Avg Deal Size"):
                ws_dashboard.merge_range(r + 1, c, r + 3, c + 3, value, fmt_kpi_value_currency)
            else:
                ws_dashboard.merge_range(r + 1, c, r + 3, c + 3, value, fmt_kpi_value_int)

        # Charts
        # Determine ranges for aggregates on Aggregates sheet
        # Find last rows
        m_rows = len(monthly_df)
        a_rows = len(by_agent_df)
        ta_rows = len(top_agents_df)
        c_rows = len(by_category_df)
        r_rows = len(by_region_df)

        # Monthly line chart
        chart_monthly = workbook.add_chart({"type": "line"})
        chart_monthly.add_series({
            "name": "=Aggregates!$B$1",
            "categories": f"=Aggregates!$A$2:$A${m_rows + 1}",
            "values": f"=Aggregates!$B$2:$B${m_rows + 1}",
        })
        chart_monthly.set_title({"name": "Monthly Revenue"})
        chart_monthly.set_y_axis({"num_format": "$#,##0"})
        chart_monthly.set_legend({"position": "none"})
        ws_dashboard.insert_chart("A12", chart_monthly, {"x_scale": 1.25, "y_scale": 1.1})

        # Top agents column chart (use top_agents_df range rather than full by_agent)
        # Top agents table is not physically written; we will anchor to ByAgent top N rows
        # ByAgent is at col E (index 4) on Aggregates: columns E and F
        chart_agents = workbook.add_chart({"type": "column"})
        # categories E2:E{1+ta_rows}, values F2:F{1+ta_rows}
        start_col_agents = start_positions["ByAgent"][1]  # 4 => column E
        cat_col = start_col_agents + 0
        val_col = start_col_agents + 1
        cat_start = 2
        cat_end = 1 + ta_rows
        chart_agents.add_series({
            "name": "Top Agents",
            "categories": f"=Aggregates!${chr(65+cat_col)}${cat_start}:${chr(65+cat_col)}${cat_end}",
            "values": f"=Aggregates!${chr(65+val_col)}${cat_start}:${chr(65+val_col)}${cat_end}",
        })
        chart_agents.set_title({"name": "Top Agents by Revenue"})
        chart_agents.set_y_axis({"num_format": "$#,##0"})
        chart_agents.set_legend({"position": "none"})
        ws_dashboard.insert_chart("E12", chart_agents, {"x_scale": 1.25, "y_scale": 1.1})

        # Category pie chart (start col I index 8)
        chart_cat = workbook.add_chart({"type": "pie"})
        start_col_cat = start_positions["ByCategory"][1]  # 8 => I
        chart_cat.add_series({
            "name": "Category Share",
            "categories": f"=Aggregates!${chr(65+start_col_cat)}$2:${chr(65+start_col_cat)}${c_rows + 1}",
            "values": f"=Aggregates!${chr(65+start_col_cat+1)}$2:${chr(65+start_col_cat+1)}${c_rows + 1}",
            "data_labels": {"percentage": True},
        })
        chart_cat.set_title({"name": "Revenue by Category"})
        ws_dashboard.insert_chart("A28", chart_cat, {"x_scale": 1.0, "y_scale": 1.0})

        # Region bar chart (start col M index 12)
        chart_reg = workbook.add_chart({"type": "bar"})
        start_col_reg = start_positions["ByRegion"][1]  # 12 => M
        chart_reg.add_series({
            "name": "Region Revenue",
            "categories": f"=Aggregates!${chr(65+start_col_reg)}$2:${chr(65+start_col_reg)}${r_rows + 1}",
            "values": f"=Aggregates!${chr(65+start_col_reg+1)}$2:${chr(65+start_col_reg+1)}${r_rows + 1}",
        })
        chart_reg.set_title({"name": "Revenue by Region"})
        chart_reg.set_y_axis({"num_format": "$#,##0"})
        chart_reg.set_legend({"position": "none"})
        ws_dashboard.insert_chart("E28", chart_reg, {"x_scale": 1.25, "y_scale": 1.0})

        # Adjust column widths on Dashboard
        ws_dashboard.set_column("A:A", 16)
        ws_dashboard.set_column("B:H", 14)

        # Tweak column widths on Aggregates and RawData using writer.sheets
        ws_agg = writer.sheets[sheet_agg]
        ws_agg.set_column("A:A", 12)  # Month
        ws_agg.set_column("B:B", 14)
        ws_agg.set_column("E:E", 18)
        ws_agg.set_column("F:F", 14)
        ws_agg.set_column("I:I", 16)
        ws_agg.set_column("J:J", 14)
        ws_agg.set_column("M:M", 14)
        ws_agg.set_column("N:N", 14)

        ws_raw = writer.sheets["RawData"]
        ws_raw.autofilter(0, 0, len(data.raw), len(data.raw.columns) - 1)
        ws_raw.freeze_panes(1, 0)
        for idx, col in enumerate(data.raw.columns):
            width = max(12, min(30, int(max(data.raw[col].astype(str).map(len).max(), len(col)) * 1.2)))
            ws_raw.set_column(idx, idx, width)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate Excel dashboard for sales agents")
    parser.add_argument("--input", type=Path, required=True, help="Path to input CSV data")
    parser.add_argument("--output", type=Path, default=Path(DEFAULT_OUTPUT_FILENAME), help="Output .xlsx path")
    parser.add_argument("--top", type=int, default=10, help="Top N agents to chart")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    df = read_sales_csv(args.input)
    data = build_aggregates(df, top_n_agents=args.top)
    output_path = args.output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    write_excel_dashboard(data, output_path)
    print(f"Dashboard written to: {output_path}")


if __name__ == "__main__":
    main()

