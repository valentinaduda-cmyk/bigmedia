"""
Second-pass grouping for an already-sorted workbook (the output of
sort_workbook() / `bigmedia sort`).

On a fixed set of category sheets (DEFAULT_SHEETS below, matched against
each sheet's title case-insensitively — pass `sheets=` to override for a
workbook that names them differently, e.g. real delivery files sometimes
use "Getty videos"/"Getty pics" instead of "Getty Videos"/"Getty Stills"
and may not have a "BBC" sheet at all), clips that are really the same
source clip are rearranged so every occurrence sits together as one
group, in ascending order by clip id. Each group gets:
  - "Total Duration" / "Seconds" columns inserted right after "Clip
    Duration" (pushing the rest of the original columns right), summing
    that group's own "Clip Duration" values at the given fps. Filled only
    on the group's LAST row, blank above it.
  - every row of the group highlighted, full row width, GREEN if its total
    duration is 4 seconds or less, YELLOW if it's more than 4 seconds —
    replacing whatever fill the source had.
  - a thick black bottom border across the full row on its last row,
    separating it from the next group.
Every other sheet (Worksheet, GFX, Getty Unknown, Artlist, Camera Footage,
3rd parties, ...) is copied through unchanged.

Two clips are "the same clip" if dedupe.dedup_key() collapses them to the
same key (extension + trailing " (1)"-style copy marker stripped) — e.g.
"GettyImages-871860158.mov" and "GettyImages-871860158 (1).mov" are the same
clip, but "GettyImages-356-30.mov" and "GettyImages-356-31.mov" are not.

Reuters is a deliberate exception: its rushes/compilation exports repeat a
tape id ("m<digits>", e.g. "..._m909021_0-20-15-0-29-40.mov") across many
clips that are each a different timecode range cut from that same tape.
Those are grouped by that tape id instead of the full dedup_key, so the
whole tape's total duration rolls up together — see _reuters_key().

"Total Seconds" and "Total Seconds for clips >4s" on each processed sheet
are live SUM/SUMIF formulas over the "Seconds" column, not hardcoded
numbers, so they stay correct if someone edits durations by hand afterward.
"""
import re
import copy
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string

from .dedupe import dedup_key
from .timecode import tc_to_frames, frames_to_tc, frames_to_seconds_ceil
from .xlsx_utils import (
    find_column_any,
    copy_sheet_verbatim,
    measure_header_widths,
    apply_header_widths,
    dominant_header_fill,
    apply_uniform_header_fill,
    apply_uniform_data_font,
    AUTOFIT_MIN_WIDTH,
    AUTOFIT_MAX_WIDTH,
)

DEFAULT_SHEETS = ["AP", "Getty Videos", "Getty Stills", "Reuters", "Shutterstock", "BBC"]

# Real delivery files sometimes use an entirely different word for the
# stills sheet, not just different casing (e.g. "Getty pics" instead of
# "Getty Stills") -- a plain case-insensitive match against the wanted
# sheet list won't catch that, so any sheet titled like a known alias is
# folded into processing too, same convention as getty_ids.py.
SHEET_ALIASES = {"getty stills": ("getty pics",)}

# Reuters rushes/compilation filenames embed a tape id as "_m<digits>_"
# (e.g. "..._m909021_0-20-15-0-29-40.mov"); different timecode ranges cut
# from the same tape must still group together. Falls back to dedup_key
# when no such id is present, so odd rows aren't silently dropped.
_REUTERS_TAPE_ID_RE = re.compile(r'_[mM](\d+)_')


def _reuters_key(name):
    match = _REUTERS_TAPE_ID_RE.search(name or "")
    return f"m{match.group(1)}" if match else dedup_key(name)


# Per-sheet override of the grouping key, matched against the sheet title
# case-insensitively; anything not listed here uses dedup_key.
SHEET_KEY_FUNCS = {"reuters": _reuters_key}


def _key_func_for_sheet(title):
    return SHEET_KEY_FUNCS.get(title.strip().lower(), dedup_key)

# Real delivery files aren't consistent about the clip-name header: some
# episodes use "Name", others "Clip Name" for the same data. Whichever one
# is passed as name_column is tried first; these cover the other case.
NAME_COLUMN_ALIASES = ("Clip Name", "Name")

# Same story for the duration header: some files use "Clip Duration",
# others just "Duration" for the same data.
DURATION_COLUMN_ALIASES = ("Clip Duration", "Duration")

THICK_BOTTOM = Border(bottom=Side(style="thick", color="000000"))
GREEN_FILL = PatternFill("solid", fgColor="92D050")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
DURATION_HIGHLIGHT_THRESHOLD = 4


def _copy_cell(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy.copy(src_cell.protection)


def _group_rows(ws_src, name_col_idx, key_func):
    """Row numbers grouped by key_func, each group's rows kept in their
    original relative order, groups themselves ordered by key ascending."""
    groups = {}
    for r in range(2, ws_src.max_row + 1):
        name = ws_src.cell(row=r, column=name_col_idx).value
        if name is None or str(name).strip() == "":
            continue
        key = key_func(name)
        groups.setdefault(key, []).append(r)
    return [groups[key] for key in sorted(groups)]


def _process_sheet(ws_src, wb_out, title, name_column, duration_column, fps):
    max_col = ws_src.max_column
    name_candidates = [name_column] + [a for a in NAME_COLUMN_ALIASES if a != name_column]
    name_col_idx = find_column_any(ws_src, name_candidates)
    duration_candidates = [duration_column] + [a for a in DURATION_COLUMN_ALIASES if a != duration_column]
    duration_col_idx = find_column_any(ws_src, duration_candidates)
    groups = _group_rows(ws_src, name_col_idx, _key_func_for_sheet(title))

    ws_out = wb_out.create_sheet(title=title)
    total_col = duration_col_idx + 1
    seconds_col = duration_col_idx + 2
    out_max_col = max_col + 2

    def out_col(src_col):
        """Original columns after "Clip Duration" shift right by two to
        make room for the inserted Total Duration/Seconds columns."""
        return src_col if src_col <= duration_col_idx else src_col + 2

    for c in range(1, max_col + 1):
        _copy_cell(ws_src.cell(row=1, column=c), ws_out.cell(row=1, column=out_col(c)))
    header_style_src = ws_src.cell(row=1, column=duration_col_idx)
    for c, label in [(total_col, "Total Duration"), (seconds_col, "Seconds")]:
        cell = ws_out.cell(row=1, column=c, value=label)
        cell.font = copy.copy(header_style_src.font)
        cell.fill = copy.copy(header_style_src.fill)
        cell.alignment = copy.copy(header_style_src.alignment)
    if 1 in ws_src.row_dimensions:
        ws_out.row_dimensions[1].height = ws_src.row_dimensions[1].height

    total_clips = 0
    total_unique = 0
    total_over_4s = 0
    total_seconds = 0
    total_seconds_over_4s = 0
    out_r = 2
    for group in groups:
        total_unique += 1
        total_clips += len(group)
        frames_sum = sum(
            tc_to_frames(ws_src.cell(row=r, column=duration_col_idx).value, fps)
            if ws_src.cell(row=r, column=duration_col_idx).value else 0
            for r in group
        )
        group_seconds = frames_to_seconds_ceil(frames_sum, fps)
        total_seconds += group_seconds
        over_4s = group_seconds > DURATION_HIGHLIGHT_THRESHOLD
        if over_4s:
            total_over_4s += 1
            total_seconds_over_4s += group_seconds
        fill = YELLOW_FILL if over_4s else GREEN_FILL

        for i, r in enumerate(group):
            for c in range(1, max_col + 1):
                _copy_cell(ws_src.cell(row=r, column=c), ws_out.cell(row=out_r, column=out_col(c)))
            if r in ws_src.row_dimensions:
                ws_out.row_dimensions[out_r].height = ws_src.row_dimensions[r].height
            if i == len(group) - 1:
                ws_out.cell(row=out_r, column=total_col, value=frames_to_tc(frames_sum, fps))
                ws_out.cell(row=out_r, column=seconds_col, value=group_seconds)
            for c in range(1, out_max_col + 1):
                cell = ws_out.cell(row=out_r, column=c)
                cell.fill = copy.copy(fill)
                if i == len(group) - 1:
                    cell.border = THICK_BOTTOM
            out_r += 1

    for src_col_letter, dim in ws_src.column_dimensions.items():
        src_col = column_index_from_string(src_col_letter)
        ws_out.column_dimensions[get_column_letter(out_col(src_col))].width = dim.width
    dur_col_letter = get_column_letter(duration_col_idx)
    dur_width = ws_src.column_dimensions[dur_col_letter].width if dur_col_letter in ws_src.column_dimensions else None
    if dur_width:
        ws_out.column_dimensions[get_column_letter(total_col)].width = dur_width
        ws_out.column_dimensions[get_column_letter(seconds_col)].width = dur_width
    ws_out.freeze_panes = "A2"

    # Seconds column only has values on each group's last row; blank rows
    # are ignored by SUM/SUMIF, so summing the whole data range works.
    seconds_col_letter = get_column_letter(seconds_col)
    seconds_range = f"{seconds_col_letter}2:{seconds_col_letter}{out_r - 1}"
    total_seconds_formula = f"=SUM({seconds_range})"
    total_seconds_over_4s_formula = f'=SUMIF({seconds_range},">{DURATION_HIGHLIGHT_THRESHOLD}")'

    summary_row = out_r + 1  # one blank spacer row after the last data row
    for i, (label, value) in enumerate([
        ("Total clips", total_clips),
        ("Total unique clips", total_unique),
        (f"Total clips >{DURATION_HIGHLIGHT_THRESHOLD}s", total_over_4s),
        ("Total Seconds", total_seconds_formula),
        (f"Total Seconds for clips >{DURATION_HIGHLIGHT_THRESHOLD}s", total_seconds_over_4s_formula),
    ]):
        r = summary_row + i
        label_cell = ws_out.cell(row=r, column=name_col_idx, value=label)
        label_cell.font = Font(bold=True)
        ws_out.cell(row=r, column=name_col_idx + 1, value=value)

    return {
        "total_clips": total_clips,
        "total_unique_clips": total_unique,
        "total_clips_over_4s": total_over_4s,
        "total_seconds": total_seconds,
        "total_seconds_over_4s": total_seconds_over_4s,
    }


def group_duplicates_workbook(
    src_path,
    out_path,
    sheets=None,
    name_column="Clip Name",
    duration_column="Clip Duration",
    fps=25,
    autofit=False,
    min_width=AUTOFIT_MIN_WIDTH,
    max_width=AUTOFIT_MAX_WIDTH,
    uniform_font=False,
    uniform_header=False,
):
    sheets = sheets or DEFAULT_SHEETS
    wanted = {s.strip().lower() for s in sheets}
    for base, aliases in SHEET_ALIASES.items():
        if base in wanted:
            wanted.update(aliases)

    wb_src = load_workbook(src_path)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    results = {}
    for title in wb_src.sheetnames:
        ws_src = wb_src[title]
        if title.strip().lower() in wanted:
            results[title] = _process_sheet(ws_src, wb_out, title, name_column, duration_column, fps)
        else:
            ws_out = wb_out.create_sheet(title=title)
            copy_sheet_verbatim(ws_src, ws_out, ws_src.max_row, ws_src.max_column)

    # Formatting runs last, over the finished workbook, because grouping
    # inserts columns as it goes -- measuring earlier would size the wrong
    # table. The "Worksheet" backup is deliberately excluded: it is a
    # verbatim copy of the input and must stay one.
    formatted = [wb_out[t] for t in wb_out.sheetnames if t != "Worksheet"]
    if autofit and formatted:
        widths = measure_header_widths(formatted, min_width=min_width, max_width=max_width)
        for ws in formatted:
            apply_header_widths(ws, widths)
    if uniform_header and formatted:
        # Measured on the source sheets, before anything is normalized, so
        # the file's own header colour wins over the fallback.
        header_fill = dominant_header_fill([wb_src[t] for t in wb_src.sheetnames if t != "Worksheet"])
        for ws in formatted:
            apply_uniform_header_fill(ws, header_fill)
    if uniform_font and formatted:
        base = _first_data_font(formatted)
        if base is not None:
            for ws in formatted:
                apply_uniform_data_font(ws, base)

    wb_out.save(out_path)
    return results


def _first_data_font(worksheets):
    """The font of the first data cell found across the given sheets, used as
    the single font every data cell is normalized to."""
    for ws in worksheets:
        for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row, 2),
                                max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    return copy.copy(cell.font)
    return None
