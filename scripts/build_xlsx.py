#!/usr/bin/env python3
"""
build_xlsx.py — Branded openpyxl workbook builder.

Reads an xlsx.json spec + report data JSON and produces a formatted .xlsx file.

Usage:
    python build_xlsx.py \
        --template rsra \
        --data /path/to/report_data.json \
        --brand /path/to/brand.json \
        --xlsx-spec /path/to/xlsx.json \
        --output /path/to/output.xlsx
"""

import argparse
import json
import os
import sys
from collections import defaultdict
from typing import Any, Dict, List, Optional

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.chart.series import DataPoint
    from openpyxl.styles import (
        Alignment,
        Border,
        Font,
        GradientFill,
        NamedStyle,
        PatternFill,
        Side,
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins, PrintPageSetup
except ImportError as e:
    print(f"ERROR: openpyxl is required. {e}", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def resolve_path(data: Any, dotpath: str) -> Any:
    """Traverse nested dict by dot-notation key. Returns None if any key missing."""
    if dotpath is None or dotpath == "":
        return None
    parts = dotpath.split(".")
    current = data
    for part in parts:
        if current is None:
            return None
        if isinstance(current, dict):
            current = current.get(part)
        else:
            return None
    return current


def hex_to_argb(hex_color: Optional[str]) -> str:
    """Convert '#1B2A3B' or '1B2A3B' to 'FF1B2A3B' for openpyxl."""
    if not hex_color:
        return "FF000000"
    h = hex_color.lstrip("#")
    if len(h) == 6:
        return f"FF{h.upper()}"
    if len(h) == 8:
        return h.upper()
    return "FF000000"


def apply_format(cell, fmt_string: Optional[str]) -> None:
    """Set cell.number_format from a format string."""
    if fmt_string:
        cell.number_format = fmt_string


def set_col_width(ws, col_idx: int, width: Optional[float]) -> None:
    """Set column width by 1-based column index."""
    if width is not None:
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = float(width)


def resolve_brand_color(value: str, brand: dict) -> str:
    """
    Resolve a color that may be a brand reference like 'brand.primary_color'
    or a hex string like '#1B2A3B'.
    """
    if value and value.startswith("brand."):
        key = value[len("brand."):]
        resolved = brand.get(key, "#000000")
        return resolved
    return value


def make_fill(hex_color: str) -> PatternFill:
    return PatternFill(
        fill_type="solid",
        fgColor=hex_to_argb(hex_color),
    )


def make_font(
    name: str = "Calibri",
    size: int = 9,
    bold: bool = False,
    color: str = "#000000",
    italic: bool = False,
) -> Font:
    return Font(
        name=name,
        size=size,
        bold=bold,
        color=hex_to_argb(color),
        italic=italic,
    )


def make_border(color: str = "#E5E7EB") -> Border:
    side = Side(style="thin", color=hex_to_argb(color))
    return Border(bottom=side)


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def build_key_value_sheet(ws, sheet_spec: dict, data: dict, brand: dict, wb_style: dict) -> None:
    """Build a key_value type sheet."""
    rows_spec = sheet_spec.get("rows", [])
    columns_spec = sheet_spec.get("columns", [])
    style_spec = sheet_spec.get("style", {})

    body_font_name = wb_style.get("body_font", brand.get("font_body", "Calibri"))
    body_font_size = wb_style.get("body_font_size", 9)
    header_fill_hex = resolve_brand_color(
        wb_style.get("header_fill_color", brand.get("primary_color", "#0F1923")), brand
    )
    header_font_color = wb_style.get("header_font_color", "#FFFFFF")
    header_font_size = wb_style.get("header_font_size", 9)
    accent_fill_hex = brand.get("accent_color", "#F0F4F8")
    border_color = brand.get("border_color", "#E5E7EB")
    alt_fill_hex = wb_style.get("alternate_row_fill", "#F9FAFB")

    # Signal level fill map from style spec
    signal_fills = {}
    signal_row_style = style_spec.get("signal_level_row", {})
    if "fill_by_value" in signal_row_style:
        signal_fills = signal_row_style["fill_by_value"]

    # Column widths
    col_widths = [c.get("width") for c in columns_spec]

    # Write header row
    headers = [c.get("header", "") for c in columns_spec]
    for ci, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=ci, value=hdr)
        cell.font = make_font(
            name=body_font_name,
            size=header_font_size,
            bold=True,
            color=header_font_color,
        )
        cell.fill = make_fill(header_fill_hex)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        set_col_width(ws, ci, col_widths[ci - 1] if ci - 1 < len(col_widths) else None)

    ws.freeze_panes = "A2"

    # Track display row (skip spacers from row counting for alternating fills, but write them)
    row_num = 2
    data_row_count = 0  # for alternate row fill

    for row_spec in rows_spec:
        row_type = row_spec.get("type")

        if row_type == "spacer":
            # Write an empty row
            ws.row_dimensions[row_num].height = 8
            row_num += 1
            continue

        label = row_spec.get("label", "")
        dotpath = row_spec.get("data_path", "")
        value = resolve_path(data, dotpath) if dotpath else None
        omit_if_null = row_spec.get("omit_if_null", False)

        if omit_if_null and value is None:
            continue

        row_style = row_spec.get("style", "")
        fmt = row_spec.get("format")
        row_height = row_spec.get("row_height")

        # Determine fills
        label_fill = make_fill(accent_fill_hex)  # default accent for label col
        value_fill = None
        row_bold = False

        if row_style == "signal_level" and value and value in signal_fills:
            color = signal_fills[value]
            label_fill = make_fill(color)
            value_fill = make_fill(color)
            row_bold = signal_row_style.get("bold", False)
        elif data_row_count % 2 == 1:
            # Alternate row tinting for value column only
            pass  # keep default fills

        # Label cell (col A)
        label_cell = ws.cell(row=row_num, column=1, value=label)
        label_cell.font = make_font(
            name=body_font_name,
            size=body_font_size,
            bold=True,
            color=brand.get("text_color", "#0F1923"),
        )
        label_cell.fill = label_fill
        label_cell.alignment = Alignment(vertical="center", wrap_text=False)
        label_cell.border = make_border(border_color)

        # Value cell (col B)
        value_cell = ws.cell(row=row_num, column=2, value=value)
        value_cell.font = make_font(
            name=body_font_name,
            size=body_font_size,
            bold=row_bold,
            color=brand.get("text_color", "#0F1923"),
        )
        if value_fill:
            value_cell.fill = value_fill
        elif data_row_count % 2 == 1:
            value_cell.fill = make_fill(alt_fill_hex)

        wrap = row_style == "wrap" or (
            len(columns_spec) > 1 and columns_spec[1].get("wrap", False)
        )
        value_cell.alignment = Alignment(vertical="top", wrap_text=wrap)
        value_cell.border = make_border(border_color)

        if fmt:
            apply_format(value_cell, fmt)

        if row_height:
            ws.row_dimensions[row_num].height = row_height

        row_num += 1
        data_row_count += 1

    # Ensure col widths for both columns
    for ci, col_spec in enumerate(columns_spec, start=1):
        set_col_width(ws, ci, col_spec.get("width"))


def build_table_sheet(ws, sheet_spec: dict, data: dict, brand: dict, wb_style: dict) -> None:
    """Build a table type sheet."""
    dotpath = sheet_spec.get("data_path", "")
    columns_spec = sheet_spec.get("columns", [])
    totals_spec = sheet_spec.get("totals_row")
    cond_fmt = sheet_spec.get("conditional_formatting", {})
    note = sheet_spec.get("note") or sheet_spec.get("footer_note")
    freeze = sheet_spec.get("freeze_panes", "A2")
    charts_spec = sheet_spec.get("charts", [])
    filter_spec = sheet_spec.get("filter")

    body_font_name = wb_style.get("body_font", brand.get("font_body", "Calibri"))
    body_font_size = wb_style.get("body_font_size", 9)
    header_fill_hex = resolve_brand_color(
        wb_style.get("header_fill_color", brand.get("primary_color", "#0F1923")), brand
    )
    header_font_color = wb_style.get("header_font_color", "#FFFFFF")
    header_font_size = wb_style.get("header_font_size", 9)
    alt_fill_hex = wb_style.get("alternate_row_fill", "#F9FAFB")
    border_color = brand.get("border_color", "#E5E7EB")
    primary_hex = resolve_brand_color(brand.get("primary_color", "#0F1923"), brand)

    # Resolve row data
    rows_data: List[dict] = resolve_path(data, dotpath) or []
    if not isinstance(rows_data, list):
        rows_data = []

    # Apply filter
    if filter_spec:
        field = filter_spec.get("field")
        not_null = filter_spec.get("not_null", False)
        if field and not_null:
            rows_data = [r for r in rows_data if r.get(field) not in (None, "", [])]

    # Determine which columns to include (omit_if_all_null)
    active_columns = []
    for col in columns_spec:
        key = col.get("data_key")
        if col.get("omit_if_all_null") and key:
            all_null = all(
                (r.get(key) is None or r.get(key) == "") for r in rows_data
            )
            if all_null:
                continue
        active_columns.append(col)

    # Header row
    for ci, col in enumerate(active_columns, start=1):
        cell = ws.cell(row=1, column=ci, value=col.get("header", ""))
        cell.font = make_font(
            name=body_font_name,
            size=header_font_size,
            bold=True,
            color=header_font_color,
        )
        cell.fill = make_fill(header_fill_hex)
        cell.alignment = Alignment(wrap_text=col.get("wrap", False), vertical="center")
        set_col_width(ws, ci, col.get("width"))

    # Auto-filter on header row
    if active_columns:
        last_col_letter = get_column_letter(len(active_columns))
        ws.auto_filter.ref = f"A1:{last_col_letter}1"

    # Freeze panes
    ws.freeze_panes = freeze

    # Data rows
    for ri, row in enumerate(rows_data, start=2):
        alt_fill = make_fill(alt_fill_hex) if (ri % 2 == 1) else None
        for ci, col in enumerate(active_columns, start=1):
            key = col.get("data_key")
            col_type = col.get("type")

            if col_type == "row_index":
                value = ri - 1
            elif key == "_value":
                # For list sheets used as table: row is a string
                value = row if isinstance(row, str) else row.get("_value", row)
            else:
                value = row.get(key) if isinstance(row, dict) else None

            cell = ws.cell(row=ri, column=ci, value=value)
            cell.font = make_font(
                name=body_font_name,
                size=body_font_size,
                color=brand.get("text_color", "#0F1923"),
            )
            cell.alignment = Alignment(
                wrap_text=col.get("wrap", False), vertical="top"
            )
            cell.border = make_border(border_color)

            if alt_fill:
                cell.fill = alt_fill

            # Conditional formatting (fill by value match)
            col_key_str = key or ""
            if col_key_str in cond_fmt and isinstance(value, str):
                rule = cond_fmt[col_key_str].get(value)
                if rule and "fill" in rule:
                    cell.fill = make_fill(rule["fill"])

            if col.get("format"):
                apply_format(cell, col["format"])

    last_data_row = 1 + len(rows_data)

    # Totals row
    if totals_spec and rows_data:
        totals_row_num = last_data_row + 1
        label_col_key = totals_spec.get("label_column")
        totals_label = totals_spec.get("label", "TOTAL")
        sum_keys = totals_spec.get("sum_columns", [])
        totals_fmt = totals_spec.get("format")
        totals_bold = totals_spec.get("bold", True)
        totals_fill_hex = resolve_brand_color(
            totals_spec.get("fill", "primary"), brand
        )
        if totals_fill_hex == "primary":
            totals_fill_hex = primary_hex

        for ci, col in enumerate(active_columns, start=1):
            key = col.get("data_key", "")
            if key == label_col_key:
                cell_value = totals_label
            elif key in sum_keys:
                # Sum numeric values
                total_val = sum(
                    (r.get(key) or 0)
                    for r in rows_data
                    if isinstance(r.get(key), (int, float))
                )
                cell_value = total_val if total_val != 0 else None
            else:
                cell_value = None

            cell = ws.cell(row=totals_row_num, column=ci, value=cell_value)
            cell.font = make_font(
                name=body_font_name,
                size=body_font_size,
                bold=totals_bold,
                color="#FFFFFF",
            )
            cell.fill = make_fill(totals_fill_hex)
            cell.alignment = Alignment(vertical="center")
            if cell_value is not None and key in sum_keys and totals_fmt:
                apply_format(cell, totals_fmt)

        last_data_row = totals_row_num

    # Note / footer note
    if note:
        note_row = last_data_row + 2
        note_cell = ws.cell(row=note_row, column=1, value=note)
        note_cell.font = Font(
            name=body_font_name,
            size=body_font_size - 1,
            italic=True,
            color=hex_to_argb(brand.get("text_muted", "#6B7280")),
        )
        note_cell.alignment = Alignment(wrap_text=True)
        # Merge across all columns for readability
        if len(active_columns) > 1:
            ws.merge_cells(
                start_row=note_row,
                start_column=1,
                end_row=note_row,
                end_column=len(active_columns),
            )
        ws.row_dimensions[note_row].height = 30

    # Charts
    for chart_spec in charts_spec:
        try:
            _build_chart(ws, chart_spec, rows_data, active_columns, brand, last_data_row)
        except Exception as e:
            print(
                f"WARNING: Chart '{chart_spec.get('id', '?')}' failed: {e}",
                file=sys.stderr,
            )


def build_list_sheet(ws, sheet_spec: dict, data: dict, brand: dict, wb_style: dict) -> None:
    """Build a list type sheet (one row per string in array)."""
    dotpath = sheet_spec.get("data_path", "")
    columns_spec = sheet_spec.get("columns", [])
    row_height = sheet_spec.get("row_height")
    header_note = sheet_spec.get("header_note")

    body_font_name = wb_style.get("body_font", brand.get("font_body", "Calibri"))
    body_font_size = wb_style.get("body_font_size", 9)
    header_fill_hex = resolve_brand_color(
        wb_style.get("header_fill_color", brand.get("primary_color", "#0F1923")), brand
    )
    header_font_color = wb_style.get("header_font_color", "#FFFFFF")
    header_font_size = wb_style.get("header_font_size", 9)
    alt_fill_hex = wb_style.get("alternate_row_fill", "#F9FAFB")
    border_color = brand.get("border_color", "#E5E7EB")

    items: List[Any] = resolve_path(data, dotpath) or []
    if not isinstance(items, list):
        items = []

    start_row = 1

    # Optional header note above the table
    if header_note:
        note_cell = ws.cell(row=1, column=1, value=header_note)
        note_cell.font = Font(
            name=body_font_name,
            size=body_font_size - 1,
            italic=True,
            color=hex_to_argb(brand.get("text_muted", "#6B7280")),
        )
        note_cell.alignment = Alignment(wrap_text=True)
        if len(columns_spec) > 1:
            ws.merge_cells(
                start_row=1,
                start_column=1,
                end_row=1,
                end_column=len(columns_spec),
            )
        ws.row_dimensions[1].height = 30
        start_row = 2

    # Header row
    header_row = start_row
    for ci, col in enumerate(columns_spec, start=1):
        cell = ws.cell(row=header_row, column=ci, value=col.get("header", ""))
        cell.font = make_font(
            name=body_font_name,
            size=header_font_size,
            bold=True,
            color=header_font_color,
        )
        cell.fill = make_fill(header_fill_hex)
        cell.alignment = Alignment(wrap_text=False, vertical="center")
        set_col_width(ws, ci, col.get("width"))

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate

    # Data rows
    for ri, item in enumerate(items, start=1):
        actual_row = header_row + ri
        alt_fill = make_fill(alt_fill_hex) if (ri % 2 == 0) else None

        for ci, col in enumerate(columns_spec, start=1):
            col_type = col.get("type")
            key = col.get("data_key", "")

            if col_type == "row_index":
                value = ri
            elif key == "_value":
                value = item if isinstance(item, str) else str(item)
            else:
                value = item.get(key, "") if isinstance(item, dict) else ""

            cell = ws.cell(row=actual_row, column=ci, value=value)
            cell.font = make_font(
                name=body_font_name,
                size=body_font_size,
                color=brand.get("text_color", "#0F1923"),
            )
            cell.alignment = Alignment(
                wrap_text=col.get("wrap", False), vertical="top"
            )
            cell.border = make_border(border_color)

            if alt_fill:
                cell.fill = alt_fill

        if row_height:
            ws.row_dimensions[actual_row].height = row_height


# ---------------------------------------------------------------------------
# Chart builder
# ---------------------------------------------------------------------------

def _build_chart(
    ws,
    chart_spec: dict,
    rows_data: List[dict],
    active_columns: List[dict],
    brand: dict,
    last_data_row: int,
) -> None:
    """Add a chart to the worksheet based on chart_spec."""
    chart_type = chart_spec.get("type", "BarChart")
    title = chart_spec.get("title", "")
    x_axis_key = chart_spec.get("x_axis")
    y_axis_key = chart_spec.get("y_axis")
    aggregate = chart_spec.get("aggregate")
    series_color_name = chart_spec.get("series_color", "primary")
    placement = chart_spec.get("placement", "right_of_table")
    width_cols = chart_spec.get("width_cols", 8)
    height_rows = chart_spec.get("height_rows", 16)
    y_axis_label = chart_spec.get("y_axis_label", "")

    # Resolve series color
    if series_color_name == "primary":
        series_hex = resolve_brand_color(brand.get("primary_color", "#0F1923"), brand)
    elif series_color_name == "highlight":
        series_hex = resolve_brand_color(brand.get("highlight_color", "#52B788"), brand)
    else:
        series_hex = resolve_brand_color(series_color_name, brand)

    series_argb = hex_to_argb(series_hex)

    # Build aggregated data if needed
    if aggregate == "sum" and x_axis_key and y_axis_key:
        agg: Dict[str, float] = defaultdict(float)
        for row in rows_data:
            x_val = row.get(x_axis_key) or "Unknown"
            y_val = row.get(y_axis_key)
            if isinstance(y_val, (int, float)):
                agg[x_val] += y_val

        keys = list(agg.keys())
        vals = [agg[k] for k in keys]
    else:
        # Use raw columns
        keys = [r.get(x_axis_key, "") for r in rows_data]
        vals = [r.get(y_axis_key, 0) for r in rows_data]

    if not keys:
        return

    # Write helper data to columns beyond the table
    helper_col_start = len(active_columns) + 2  # leave a gap column
    helper_row_start = 2  # row 1 will be the header

    ws.cell(row=1, column=helper_col_start, value=x_axis_key or "Category")
    ws.cell(row=1, column=helper_col_start + 1, value=y_axis_key or "Value")

    for i, (k, v) in enumerate(zip(keys, vals)):
        ws.cell(row=helper_row_start + i, column=helper_col_start, value=k)
        ws.cell(row=helper_row_start + i, column=helper_col_start + 1, value=v)

    helper_rows = len(keys)

    cats = Reference(
        ws,
        min_col=helper_col_start,
        min_row=helper_row_start,
        max_row=helper_row_start + helper_rows - 1,
    )
    data_ref = Reference(
        ws,
        min_col=helper_col_start + 1,
        min_row=1,  # include header as series title
        max_row=helper_row_start + helper_rows - 1,
    )

    # Create chart object
    if chart_type == "BarChart":
        chart = BarChart()
        chart.type = "bar"
        chart.grouping = "clustered"
    elif chart_type == "LineChart":
        chart = LineChart()
    elif chart_type == "PieChart":
        chart = PieChart()
    else:
        chart = BarChart()
        chart.type = "col"
        chart.grouping = "clustered"

    chart.title = title
    if y_axis_label and hasattr(chart, "y_axis"):
        chart.y_axis.title = y_axis_label

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)

    # Apply series color
    if chart.series:
        ser = chart.series[0]
        try:
            from openpyxl.drawing.fill import ColorChoice
            from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
        except ImportError:
            pass

        try:
            ser.graphicalProperties.solidFill = series_argb
        except Exception:
            pass

    # Chart size: approximate conversion — each col ~72px wide, each row ~15px tall
    chart.width = width_cols * 1.8  # cm
    chart.height = height_rows * 0.5  # cm

    # Placement anchor
    if placement == "right_of_table":
        anchor_col = len(active_columns) + 1
        anchor = f"{get_column_letter(anchor_col)}1"
    else:
        anchor = placement  # treat as literal cell address

    ws.add_chart(chart, anchor)


# ---------------------------------------------------------------------------
# Workbook builder
# ---------------------------------------------------------------------------

def build_workbook(
    report_data: dict,
    brand: dict,
    xlsx_spec: dict,
) -> Workbook:
    wb = Workbook()
    # Remove default sheet
    if wb.worksheets:
        wb.remove(wb.active)

    wb_style = xlsx_spec.get("workbook_style", {})
    sheets_spec = xlsx_spec.get("sheets", [])

    # Sort sheets by order
    sheets_spec_sorted = sorted(sheets_spec, key=lambda s: s.get("order", 999))

    # Resolve tab colors (brand references)
    raw_tab_colors = wb_style.get("tab_colors", {})
    tab_colors = {
        name: resolve_brand_color(color, brand)
        for name, color in raw_tab_colors.items()
    }

    for sheet_spec in sheets_spec_sorted:
        sheet_id = sheet_spec.get("id", "sheet")
        sheet_name = sheet_spec.get("name", sheet_id)
        sheet_type = sheet_spec.get("type", "table")
        omit_if_empty = sheet_spec.get("omit_if_empty", False)
        omit_if_path_absent = sheet_spec.get("omit_if_path_absent")

        # Check omit_if_path_absent
        if omit_if_path_absent:
            check_val = resolve_path(report_data, omit_if_path_absent)
            if check_val is None:
                continue

        # Check omit_if_empty for list/table sheets
        if omit_if_empty:
            dotpath = sheet_spec.get("data_path", "")
            arr = resolve_path(report_data, dotpath)
            if not arr:
                continue

        ws = wb.create_sheet(title=sheet_name)

        # Apply tab color
        if sheet_name in tab_colors:
            ws.sheet_properties.tabColor = hex_to_argb(tab_colors[sheet_name])

        # Print settings
        print_settings = wb_style.get("print_settings", {})
        if print_settings.get("orientation") == "landscape":
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        else:
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT

        if print_settings.get("fit_to_page"):
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0  # auto height

        if print_settings.get("repeat_header_rows"):
            ws.print_title_rows = "1:1"

        # Build sheet content by type
        if sheet_type == "key_value":
            build_key_value_sheet(ws, sheet_spec, report_data, brand, wb_style)
        elif sheet_type == "table":
            build_table_sheet(ws, sheet_spec, report_data, brand, wb_style)
        elif sheet_type == "list":
            build_list_sheet(ws, sheet_spec, report_data, brand, wb_style)
        else:
            print(
                f"WARNING: Unknown sheet type '{sheet_type}' for sheet '{sheet_name}'",
                file=sys.stderr,
            )

    return wb


# ---------------------------------------------------------------------------
# CLI entrypoint
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Build a branded .xlsx report from a spec and data JSON."
    )
    parser.add_argument("--template", help="Template name (informational only)")
    parser.add_argument("--data", required=True, help="Path to report data JSON")
    parser.add_argument("--brand", required=True, help="Path to brand JSON")
    parser.add_argument("--xlsx-spec", required=True, dest="xlsx_spec",
                        help="Path to xlsx.json spec")
    parser.add_argument("--output", required=True, help="Output .xlsx path")
    return parser.parse_args()


def load_json(path: str, label: str) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"ERROR: {label} file not found: {path}", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"ERROR: {label} JSON is invalid: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    args = parse_args()

    report_data = load_json(args.data, "report data")
    brand = load_json(args.brand, "brand")
    xlsx_spec = load_json(args.xlsx_spec, "xlsx spec")

    try:
        wb = build_workbook(report_data, brand, xlsx_spec)
    except Exception as e:
        print(f"ERROR: Failed to build workbook: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)

    if not wb.worksheets:
        print("WARNING: No sheets were generated — check data and spec.", file=sys.stderr)

    output_path = os.path.abspath(args.output)
    try:
        wb.save(output_path)
    except Exception as e:
        print(f"ERROR: Failed to save workbook to {output_path}: {e}", file=sys.stderr)
        sys.exit(1)

    print(output_path)


if __name__ == "__main__":
    main()
