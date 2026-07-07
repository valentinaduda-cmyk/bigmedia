"""
Duplicate detection for clip names.

dedup_key() is the original, proven logic (strips extension and trailing
"(1)"-style copy markers so "Clip_A.mov" and "Clip_A (1).mxf" collapse to
the same key). It also swallows a trailing " <digits>" annotation right
after the extension (e.g. a stray frame-rate/version tag), so
"GettyImages-627-122.mov" and "GettyImages-627-122.mov 25" collapse to the
same key too.

dedupe_workbook() is a first-pass CLI wrapper around it: it groups rows by
key, keeps the first occurrence, and writes a "Duplicates" sheet listing
everything that was dropped so nothing disappears silently. Treat this as
a starting point — refine the "which row wins when there's a duplicate"
rule together once we've run it against a real file with real duplicates.

Every row's "Source Reel Name" column is overwritten with the name of the
sheet it ended up on ("Deduped" or "Duplicates"), including on the
"Worksheet" backup tab.
"""
import re
import copy
from openpyxl import load_workbook, Workbook

from .xlsx_utils import find_column, capture_header_template, write_header_row, write_data_row, copy_sheet_verbatim


def dedup_key(name: str) -> str:
    base = (name or "").strip()
    base = re.sub(r'\.[A-Za-z0-9]{2,4}(?:\s+\d+)?$', '', base)  # strip extension + optional trailing " 25"-style tag
    base = re.sub(r'\s*\(\d+\)\s*$', '', base).strip()          # strip " (1)" copy marker
    return base


def dedupe_workbook(src_path, out_path, name_column="Clip Name"):
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

    seen = {}
    kept_rows = []
    dropped_rows = []
    for r in range(2, max_row + 1):
        name = ws_src.cell(row=r, column=name_col_idx).value
        if name is None or str(name).strip() == "":
            continue
        key = dedup_key(name)
        if key in seen:
            dropped_rows.append(r)
        else:
            seen[key] = r
            kept_rows.append(r)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    ws_backup = wb_out.create_sheet(title="Worksheet")
    copy_sheet_verbatim(ws_src, ws_backup, max_row, max_col)
    for r in kept_rows:
        ws_backup.cell(row=r, column=source_col_idx, value="Deduped")
    for r in dropped_rows:
        ws_backup.cell(row=r, column=source_col_idx, value="Duplicates")

    for title, rows in [("Deduped", kept_rows), ("Duplicates", dropped_rows)]:
        ws_out = wb_out.create_sheet(title=title)
        write_header_row(ws_out, header_cells, ws_src.row_dimensions[1].height)
        for out_r, src_r in enumerate(rows, start=2):
            fill = fill_even if out_r % 2 == 0 else fill_odd
            values = [ws_src.cell(row=src_r, column=c).value for c in range(1, max_col + 1)]
            values[source_col_idx - 1] = title
            write_data_row(ws_out, out_r, values, col_font, col_numfmt, fill)
        for col_letter, width in col_widths.items():
            ws_out.column_dimensions[col_letter].width = width
        ws_out.freeze_panes = "A2"

    wb_out.save(out_path)
    return {"kept": len(kept_rows), "dropped": len(dropped_rows)}
