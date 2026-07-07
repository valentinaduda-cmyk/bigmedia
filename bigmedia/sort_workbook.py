"""
Sort a media master list into per-source sheets, plus a backup.

Output layout:
  - "Worksheet"        copy of the input, with "Source Reel Name" filled in
                        per row to whichever category that row was sorted to
  - one sheet per category in CATEGORY_ORDER, e.g. "AP", "Getty Videos", ...
                        each with its own "Source Reel Name" column filled
                        with that sheet's own category name

The name column defaults to "Clip Name" — pass name_column= to override for
a workbook that uses a different header. The workbook is expected to already
have a "Source Reel Name" column; its existing values are overwritten.
"""
import copy
from openpyxl import load_workbook, Workbook

from .classify import classify, CATEGORY_ORDER
from .xlsx_utils import (
    find_column,
    capture_header_template,
    write_header_row,
    write_data_row,
    copy_sheet_verbatim,
)


def sort_workbook(src_path, out_path, name_column="Clip Name", category_order=None):
    category_order = category_order or CATEGORY_ORDER

    wb_src = load_workbook(src_path)
    ws_src = wb_src.active
    max_col = ws_src.max_column
    max_row = ws_src.max_row

    template = capture_header_template(ws_src)
    header_cells = template["header_cells"]
    col_font = template["col_font"]
    col_numfmt = template["col_numfmt"]
    col_widths = template["col_widths"]

    fill_even = copy.copy(ws_src.cell(row=2, column=2).fill)
    fill_odd = copy.copy(ws_src.cell(row=3, column=2).fill)

    name_col_idx = find_column(ws_src, name_column)
    source_col_idx = find_column(ws_src, "Source Reel Name")

    rows_by_cat = {c: [] for c in category_order}
    row_category = {}
    for r in range(2, max_row + 1):
        name = ws_src.cell(row=r, column=name_col_idx).value
        if name is None or str(name).strip() == "":
            continue
        cat = classify(name)
        rows_by_cat.setdefault(cat, [])
        rows_by_cat[cat].append(r)
        row_category[r] = cat

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    # "Worksheet" backup tab always goes first.
    ws_backup = wb_out.create_sheet(title="Worksheet")
    copy_sheet_verbatim(ws_src, ws_backup, max_row, max_col)
    for r, cat in row_category.items():
        ws_backup.cell(row=r, column=source_col_idx, value=cat)

    for cat in category_order:
        ws_out = wb_out.create_sheet(title=cat[:31])
        write_header_row(ws_out, header_cells, ws_src.row_dimensions[1].height)

        for out_r, src_r in enumerate(rows_by_cat.get(cat, []), start=2):
            fill = fill_even if out_r % 2 == 0 else fill_odd
            values = [ws_src.cell(row=src_r, column=c).value for c in range(1, max_col + 1)]
            values[source_col_idx - 1] = cat
            write_data_row(ws_out, out_r, values, col_font, col_numfmt, fill)

        for col_letter, width in col_widths.items():
            ws_out.column_dimensions[col_letter].width = width
        ws_out.freeze_panes = "A2"

    wb_out.save(out_path)
    return {cat: len(rows) for cat, rows in rows_by_cat.items()}
