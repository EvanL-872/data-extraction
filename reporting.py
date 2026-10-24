"""PowerBI/Tableau-ready summaries and an in-memory Excel export."""

from __future__ import annotations

import io
from typing import Any

import pandas as pd


def _sum_or_blank(values: pd.Series) -> float | None:
    numeric = pd.to_numeric(values, errors="coerce").dropna()
    return float(numeric.sum()) if not numeric.empty else None


def _empty(columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columns)


def analysis_frames(frames: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Create aggregate tables without summing unlike currencies together."""
    items = frames["items"].copy()
    issues = frames["issues"].copy()
    project_columns = [
        "file", "sheet", "customer", "sales_rep", "sell_currency", "item_count",
        "hardware_item_count", "service_item_count", "total_sell", "total_cost",
        "total_sell_sgd", "total_cost_sgd", "gross_profit_sgd", "warning_count",
    ]
    customer_columns = ["customer", "sell_currency", "project_count", "total_sell", "total_sell_sgd"]
    category_columns = [
        "brand_category", "segment", "sell_currency", "item_count", "total_sell",
        "total_sell_sgd", "total_cost_sgd", "gross_profit_sgd",
    ]
    if items.empty:
        return {
            "project_summary": _empty(project_columns),
            "customer_revenue": _empty(customer_columns),
            "category_summary": _empty(category_columns),
        }

    group_fields = ["file", "sheet", "customer", "sales_rep", "sell_currency"]
    project = (
        items.groupby(group_fields, dropna=False)
        .agg(
            item_count=("item_name", "size"),
            hardware_item_count=("segment", lambda values: int((values == "Hardware / Items").sum())),
            service_item_count=("segment", lambda values: int((values == "Professional Services").sum())),
            total_sell=("total_sell", _sum_or_blank),
            total_cost=("total_cost", _sum_or_blank),
            total_sell_sgd=("total_sell_sgd", _sum_or_blank),
            total_cost_sgd=("total_cost_sgd", _sum_or_blank),
        )
        .reset_index()
    )
    project["gross_profit_sgd"] = project["total_sell_sgd"] - project["total_cost_sgd"]
    warnings = (
        issues.groupby(["file", "sheet"], dropna=False)
        .size()
        .rename("warning_count")
        .reset_index()
    )
    project = project.merge(warnings, on=["file", "sheet"], how="left")
    project["warning_count"] = project["warning_count"].fillna(0).astype(int)
    project = project[project_columns]

    customer = (
        project.groupby(["customer", "sell_currency"], dropna=False)
        .agg(
            project_count=("file", "nunique"),
            total_sell=("total_sell", _sum_or_blank),
            total_sell_sgd=("total_sell_sgd", _sum_or_blank),
        )
        .reset_index()
    )
    category = (
        items.groupby(["brand_category", "segment", "sell_currency"], dropna=False)
        .agg(
            item_count=("item_name", "size"),
            total_sell=("total_sell", _sum_or_blank),
            total_sell_sgd=("total_sell_sgd", _sum_or_blank),
            total_cost_sgd=("total_cost_sgd", _sum_or_blank),
        )
        .reset_index()
    )
    category["gross_profit_sgd"] = category["total_sell_sgd"] - category["total_cost_sgd"]
    return {
        "project_summary": project,
        "customer_revenue": customer[customer_columns],
        "category_summary": category[category_columns],
    }


def _write_frame(
    writer: pd.ExcelWriter,
    name: str,
    frame: pd.DataFrame,
    currency_columns: set[str],
) -> None:
    safe = frame.copy()
    safe.to_excel(writer, sheet_name=name, index=False, freeze_panes=(1, 0))
    workbook = writer.book
    worksheet = writer.sheets[name]
    header = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "border": 0})
    money = workbook.add_format({"num_format": "#,##0.00;[Red]-#,##0.00;–"})
    integer = workbook.add_format({"num_format": "0"})
    for index, column in enumerate(safe.columns):
        values = safe[column].fillna("").astype(str)
        width = min(max(len(str(column)) + 2, values.map(len).head(1000).max() + 2 if not values.empty else 0), 45)
        cell_format = money if column in currency_columns else integer if column.endswith("_count") or column in {"excel_row", "start_row", "end_row"} else None
        worksheet.set_column(index, index, width, cell_format)
        worksheet.write(0, index, column, header)
    if len(safe.columns) and not safe.empty:
        worksheet.autofilter(0, 0, len(safe), len(safe.columns) - 1)


def build_excel_export(frames: dict[str, pd.DataFrame]) -> bytes:
    """Build the complete workbook in memory; neither source nor export is written to disk."""
    analyses = analysis_frames(frames)
    info = pd.DataFrame([
        {"Worksheet": "All Items", "Purpose": "One grouped item per row; the primary table for BI tools."},
        {"Worksheet": "Hardware Items", "Purpose": "All non-professional-service items, including licenses, support and discounts."},
        {"Worksheet": "Professional Services", "Purpose": "Professional Service items separated from hardware/items."},
        {"Worksheet": "Project Summary", "Purpose": "Project-level sales, cost and gross profit. Native currency is never aggregated across currencies."},
        {"Worksheet": "Customer Revenue", "Purpose": "Customer revenue across the uploaded projects, grouped by selling currency."},
        {"Worksheet": "Category Summary", "Purpose": "Brand/category analysis (e.g. Ruckus, Aruba, Logitech) grouped by currency."},
        {"Worksheet": "Review Issues", "Purpose": "Rows needing review, including missing prices and reconciliation variances."},
        {"Worksheet": "Source Lines", "Purpose": "Auditable source rows behind every grouped item."},
    ])
    items = frames["items"]
    presentation_items = items.drop(columns=["block_id", "sn", "start_row", "end_row"], errors="ignore")
    tables: list[tuple[str, pd.DataFrame]] = [
        ("Read Me", info),
        ("All Items", items),
        ("Hardware Items", presentation_items[presentation_items["segment"] == "Hardware / Items"].copy()),
        ("Professional Services", presentation_items[presentation_items["segment"] == "Professional Services"].copy()),
        ("Project Summary", analyses["project_summary"]),
        ("Customer Revenue", analyses["customer_revenue"]),
        ("Category Summary", analyses["category_summary"]),
        ("Suppliers", frames["suppliers"]),
        ("Source Lines", frames["lines"]),
        ("Review Issues", frames["issues"]),
        ("Adjustments", frames["adjustments"]),
        ("Reconciliation", frames["reconciliation"]),
    ]
    money_columns = {
        "total_cost", "total_sell", "total_cost_sgd", "total_sell_sgd", "declared_cost",
        "cost_variance", "cost", "sell", "recorded_sgd", "pr_summary", "variance",
    }
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, frame in tables:
            _write_frame(writer, name, frame, money_columns)
    return output.getvalue()
