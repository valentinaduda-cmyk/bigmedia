"""
Shared openpyxl helpers used by every workbook-processing command.

Keeping these here means sort_workbook.py, dedupe.py, and any future
command (report generation, QC checks, ...) don't each reinvent "copy this
sheet with its styling" from scratch.
"""
import copy
from pathlib import Path
from openpyxl.utils import get_column_letter


def iter_xlsx_files(path, exclude_suffix=None):
    """Path objects for the .xlsx file(s) at `path`. If `path` is a folder,
    every .xlsx file directly inside it is included, except Excel's own
    "~$foo.xlsx" lock files for currently-open workbooks. If `exclude_suffix`
    is given (e.g. "grouped"), files whose stem ends with "_<exclude_suffix>"
    are also skipped -- pass the current command's own output suffix so a
    batch command re-run on the same folder doesn't reprocess its own prior
    output every time (which would otherwise snowball file count on each
    run: "x.xlsx" -> "x_grouped.xlsx" -> "x_grouped_grouped.xlsx" -> ...).
    Other stages' suffixes (e.g. "_sorted" files being fed into "group") are
    legitimate input and must NOT be filtered out. If `path` is a single
    file, it's returned as a one-item list unfiltered."""
    p = Path(path)
    if p.is_dir():
        files = (f for f in p.glob("*.xlsx") if not f.name.startswith("~$"))
        if exclude_suffix:
            files = (f for f in files if not f.stem.endswith(f"_{exclude_suffix}"))
        return sorted(files)
    return [p]


def find_column(ws, header_name: str) -> int:
    """1-indexed column number of the first header cell matching header_name
    (case-sensitive, exact match) in row 1. Raises ValueError if not found."""
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=c).value == header_name:
            return c
    raise ValueError(f"Column {header_name!r} not found in header row")


def find_column_any(ws, header_names) -> int:
    """Like find_column, but tries each name in header_names in order and
    returns the first match. Needed because real delivery files aren't
    consistent about column naming (e.g. some use "Name", others "Clip
    Name" for the same data). Raises ValueError if none match."""
    for header_name in header_names:
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=1, column=c).value == header_name:
                return c
    raise ValueError(f"None of {list(header_names)!r} found in header row")


def capture_header_template(ws):
    """Snapshot everything needed to reproduce this sheet's header row and
    per-column styling elsewhere: header cells, a font/number-format sample
    from row 2, and configured column widths."""
    max_col = ws.max_column
    header_cells = [ws.cell(row=1, column=c) for c in range(1, max_col + 1)]
    col_font = [copy.copy(ws.cell(row=2, column=c).font) for c in range(1, max_col + 1)]
    col_numfmt = [ws.cell(row=2, column=c).number_format for c in range(1, max_col + 1)]
    col_widths = {
        get_column_letter(c): ws.column_dimensions[get_column_letter(c)].width
        for c in range(1, max_col + 1)
        if get_column_letter(c) in ws.column_dimensions
    }
    return {
        "header_cells": header_cells,
        "col_font": col_font,
        "col_numfmt": col_numfmt,
        "col_widths": col_widths,
        "max_col": max_col,
    }


def write_header_row(ws_out, header_cells, row_height=None):
    """Write a styled header row (row 1) onto ws_out from captured header cells."""
    for c, hcell in enumerate(header_cells, start=1):
        nc = ws_out.cell(row=1, column=c, value=hcell.value)
        nc.font = copy.copy(hcell.font)
        nc.fill = copy.copy(hcell.fill)
        nc.alignment = copy.copy(hcell.alignment)
        nc.number_format = hcell.number_format
    if row_height is not None:
        ws_out.row_dimensions[1].height = row_height


def write_data_row(ws_out, out_row, values, col_font, col_numfmt, fill):
    """Write one data row with the captured per-column font/number-format
    and a single fill (e.g. alternating zebra-striping) applied to every cell."""
    for c, val in enumerate(values, start=1):
        nc = ws_out.cell(row=out_row, column=c, value=val)
        nc.font = col_font[c - 1]
        nc.number_format = col_numfmt[c - 1]
        nc.fill = fill


def copy_sheet_verbatim(src_ws, dst_ws, max_row=None, max_col=None):
    """Clone src_ws into dst_ws exactly as-is: values, per-cell styling, row
    heights, column widths, merged cells, and freeze panes. Used for
    untouched backup copies of raw input data."""
    max_row = max_row or src_ws.max_row
    max_col = max_col or src_ws.max_column

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            src_cell = src_ws.cell(row=r, column=c)
            dst_cell = dst_ws.cell(row=r, column=c, value=src_cell.value)
            if src_cell.has_style:
                dst_cell.font = copy.copy(src_cell.font)
                dst_cell.fill = copy.copy(src_cell.fill)
                dst_cell.border = copy.copy(src_cell.border)
                dst_cell.alignment = copy.copy(src_cell.alignment)
                dst_cell.number_format = src_cell.number_format
                dst_cell.protection = copy.copy(src_cell.protection)
        if r in src_ws.row_dimensions:
            dst_ws.row_dimensions[r].height = src_ws.row_dimensions[r].height

    for col_letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = dim.width

    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

    dst_ws.freeze_panes = src_ws.freeze_panes
