"""
FU grid: the legal follow-up sheet built from an already-sorted workbook's
"3rd parties" sheet (the output of `bigmedia sort`).

Every clip on that sheet gets one row, carrying how many times it is used
in the cut and the summed duration of all those uses, so the legal team
can chase one clearance per clip instead of per edit. The remaining
columns (SCREENSHOTS, URL LINK, PREVIOUS LEGAL CHECK, SOURCE, FINAL FU
LEGAL NOTE, PREVIEW:) are written empty for them to fill in by hand.

Two rows are the same clip only if their "Clip Name" matches EXACTLY --
deliberately NOT dedupe.dedup_key(). This grid is a clearance checklist,
so two files whose names differ at all (even only by extension, e.g.
"...WTC.mp4" vs "...WTC.mov") are chased separately. Don't "fix" this to
reuse dedup_key.

Within a group the row with the earliest "Sequence In" wins: its TC
IN/OUT, clip duration, source duration and clip name are what land on the
grid, never a mix of several rows. Grid rows are ordered by that TC IN.
"TOTAL DURATION if multiple uses" is always filled in despite its name --
for a single-use clip it just repeats that clip's own duration, which the
team prefers over a blank cell. Only "Clip Duration" is ever parsed as timecode
-- real files carry "Source Duration" values at a different frame rate
(e.g. "00:00:04:42"), so that column is copied through as text.

Styling (colors, widths, the trailing space in the "PREVIEW: " header) is
cloned from the hand-made example grid the team already works with.
"""
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .timecode import tc_to_frames, frames_to_tc
from .xlsx_utils import find_column_any, copy_sheet_verbatim

# Real delivery files aren't consistent about these headers -- same
# aliases the other commands accept.
NAME_COLUMN_ALIASES = ("Clip Name", "Name")
DURATION_COLUMN_ALIASES = ("Clip Duration", "Duration")

GRID_SHEET_TITLE = "FU grid"

# Trailing space on "PREVIEW: " is verbatim from the example grid.
HEADERS = [
    "TC IN", "TC OUT", "CLIP DURATION", "TOTAL USES",
    "TOTAL DURATION if multiple uses", "SOURCE DURATION", "SCREENSHOTS",
    "URL LINK", "PREVIOUS LEGAL CHECK", "SOURCE", "CLIP NAME",
    "FINAL FU LEGAL NOTE", "PREVIEW: ",
]

LEGAL_CHECK_COL = HEADERS.index("PREVIOUS LEGAL CHECK") + 1
URL_COL = HEADERS.index("URL LINK") + 1
BOLD_COLS = (
    HEADERS.index("TOTAL USES") + 1,
    HEADERS.index("TOTAL DURATION if multiple uses") + 1,
    HEADERS.index("SOURCE DURATION") + 1,
)

HEADER_FILL = PatternFill("solid", fgColor="FFCFE2F3")
DATA_FILL = PatternFill("solid", fgColor="FFF4CCCC")
LEGAL_CHECK_FILL = PatternFill("solid", fgColor="FFD9EAD3")
HEADER_BORDER = Border(*(Side(style="thin"),) * 4)
CENTERED = Alignment(horizontal="center", vertical="center", wrap_text=True)
BASE_FONT = Font(name="Calibri", size=11)
BOLD_FONT = Font(name="Calibri", size=11, bold=True)
URL_FONT = Font(name="Calibri", size=11, color="FF0000FF")
HEADER_ROW_HEIGHT = 43.2

# Verbatim from the example grid; F and G are left at the default width
# there, so they're left out here too.
COLUMN_WIDTHS = {
    "A": 15.5546875, "B": 14.5546875, "C": 14.44140625, "D": 12.5546875,
    "E": 18.33203125, "H": 32.0, "I": 12.109375, "J": 11.5546875,
    "K": 77.88671875, "L": 14.5546875, "M": 12.0, "N": 8.6640625,
}


def _find_sheet(wb, sheet):
    for title in wb.sheetnames:
        if title.strip().lower() == sheet.strip().lower():
            return wb[title]
    raise ValueError(f"Sheet {sheet!r} not found; workbook has {wb.sheetnames!r}")


def _frames(value, fps):
    """Timecode cell as frames, treating blank/unparseable values as 0 --
    a stray malformed cell must not abort a whole delivery's grid."""
    try:
        return tc_to_frames(str(value), fps)
    except (AttributeError, ValueError):
        return 0


def _group_rows(ws, name_col_idx, seq_in_col_idx, fps):
    """Rows grouped by exact clip name. Each group is (surviving row, all
    its rows); groups are ordered by the surviving row's Sequence In."""
    groups = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=name_col_idx).value
        if name is None or str(name).strip() == "":
            continue
        groups.setdefault(str(name), []).append(r)

    def seq_in_frames(row):
        return _frames(ws.cell(row=row, column=seq_in_col_idx).value, fps)

    survivors = [(min(rows, key=seq_in_frames), rows) for rows in groups.values()]
    return sorted(survivors, key=lambda pair: seq_in_frames(pair[0]))


def _write_grid(ws_src, wb_out, groups, columns, fps):
    ws_out = wb_out.create_sheet(title=GRID_SHEET_TITLE)

    for c, header in enumerate(HEADERS, start=1):
        cell = ws_out.cell(row=1, column=c, value=header)
        cell.font = BOLD_FONT
        cell.fill = HEADER_FILL
        cell.border = HEADER_BORDER
        cell.alignment = CENTERED
    ws_out.row_dimensions[1].height = HEADER_ROW_HEIGHT

    for out_r, (survivor, rows) in enumerate(groups, start=2):
        def source(col_idx):
            return ws_src.cell(row=survivor, column=col_idx).value

        total_frames = sum(_frames(ws_src.cell(row=r, column=columns["duration"]).value, fps) for r in rows)
        values = [
            source(columns["seq_in"]),
            source(columns["seq_out"]),
            source(columns["duration"]),
            len(rows),
            frames_to_tc(total_frames, fps),
            source(columns["source_duration"]),
            None, None, None, None,
            source(columns["name"]),
            None, None,
        ]
        for c, value in enumerate(values, start=1):
            cell = ws_out.cell(row=out_r, column=c, value=value)
            cell.font = BOLD_FONT if c in BOLD_COLS else URL_FONT if c == URL_COL else BASE_FONT
            cell.fill = LEGAL_CHECK_FILL if c == LEGAL_CHECK_COL else DATA_FILL
            cell.alignment = CENTERED

    for col_letter, width in COLUMN_WIDTHS.items():
        ws_out.column_dimensions[col_letter].width = width
    return ws_out


def fu_grid_workbook(
    src_path,
    out_path,
    sheet="3rd parties",
    name_column="Clip Name",
    duration_column="Clip Duration",
    fps=25,
):
    wb_src = load_workbook(src_path)
    ws_src = _find_sheet(wb_src, sheet)

    columns = {
        "name": find_column_any(ws_src, [name_column] + [a for a in NAME_COLUMN_ALIASES if a != name_column]),
        "duration": find_column_any(ws_src, [duration_column] + [a for a in DURATION_COLUMN_ALIASES if a != duration_column]),
        "seq_in": find_column_any(ws_src, ["Sequence In"]),
        "seq_out": find_column_any(ws_src, ["Sequence Out"]),
        "source_duration": find_column_any(ws_src, ["Source Duration"]),
    }
    groups = _group_rows(ws_src, columns["name"], columns["seq_in"], fps)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    ws_backup = wb_out.create_sheet(title=ws_src.title)
    copy_sheet_verbatim(ws_src, ws_backup)
    _write_grid(ws_src, wb_out, groups, columns, fps)
    wb_out.save(out_path)

    return {
        "rows_in": sum(len(rows) for _, rows in groups),
        "unique_clips": len(groups),
        "multi_use_clips": sum(1 for _, rows in groups if len(rows) > 1),
    }
