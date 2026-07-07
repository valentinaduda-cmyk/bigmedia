"""
Extract Getty clip ids from clip filenames and consolidate them from
several already-sorted workbooks into one new report file, styled to match
Getty's own "Customer Declaration Form" template so it can be pasted/sent
back to Getty with no manual reformatting.

Reads the "Getty videos" sheet of every input workbook, extracts each
clip's Getty id from its filename (see extract_getty_id() for the exact
rules), and writes one new workbook with one two-column block per input
file, left to right in input order, separated by a blank spacer column:

Only UNIQUE clips with a total duration of 5+ seconds are included. Real
delivery files (and our own `group` command's output) only populate the
"Seconds" column once per clip — on the last row of a duplicate run, blank
on the earlier rows of that same run — so filtering to "Seconds present
and >= min_seconds" naturally selects one row per unique clip. An explicit
id-based dedup is layered on top as a safety net in case a file doesn't
follow that convention.

    <episode title>                     <blank>   <episode title>
    Getty Images Video                            Getty Images Video
    Asset ID | Duration                            Asset ID | Duration
    =COUNTA  | =SUM                                =COUNTA  | =SUM
    <id>     | <seconds>                           <id>     | <seconds>
    ...
"""
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .xlsx_utils import find_column, iter_xlsx_files

_GETTY_PREFIX_RE = re.compile(r'^GettyImages-', re.I)
_MR_TAG_RE = re.compile(r'^mr_', re.I)
_EXT_RE = re.compile(r'\.(mov|mp4|jpg)\b', re.I)
_KOPIE_SUFFIX_RE = re.compile(r'\s*\(kopie\)\s*$', re.I)


def extract_getty_id(name: str) -> str:
    """
    Pull the Getty id out of a clip filename. Rules, derived from real
    examples (see tests/test_getty_ids.py for the full regression list):
      - strip a leading "GettyImages-"/"GETTYIMAGES-" prefix
      - strip a leading "mr_" tag right after that prefix (Getty metadata,
        not part of the id) — e.g. "mr_00108323.mov" -> "00108323"
      - if a .mov/.mp4/.jpg extension appears anywhere, cut the string at the
        start of that extension, dropping it and everything after it
        (including trailing " 25"-style suffixes). Within what's left,
        drop a trailing chain of "_word" suffixes (e.g. "_Apple_ProRes_422",
        "_Denoise_02") UNLESS the suffix is purely numeric (e.g. "_0003"),
        which is kept as part of the id.
      - if no extension is present, keep the remainder as-is (including any
        trailing " 2"-style suffix), except a bare trailing hyphen with
        nothing after it (e.g. "96410456-" -> "96410456").
    """
    s = (name or "").strip()
    s = _GETTY_PREFIX_RE.sub('', s)
    s = _MR_TAG_RE.sub('', s)

    m = _EXT_RE.search(s)
    if m:
        s = s[:m.start()]
        parts = s.split('_')
        core = [parts[0]]
        for part in parts[1:]:
            if part.isdigit():
                core.append(part)
            else:
                break
        s = '_'.join(core)
    else:
        s = re.sub(r'-$', '', s)
    return s


def clean_episode_title(filename: str) -> str:
    """
    Block-header title for a file, e.g. "EP3 - Versailles (kopie).xlsx" ->
    "EP3 - Versailles". Strips the extension and a trailing "(kopie)" tag
    real delivery files carry from being duplicated in OneDrive/Explorer.
    """
    stem = Path(filename).stem
    return _KOPIE_SUFFIX_RE.sub('', stem).strip()


_NAME_COLUMN_CANDIDATES = ["Clip Name", "Name"]


def _find_name_column(ws, name_column):
    """Real delivery files disagree on the clip-name header ("Clip Name" on
    some, plain "Name" on others) even within the same batch, so try the
    caller's preferred name first, then fall back through known aliases."""
    candidates = [name_column] + [c for c in _NAME_COLUMN_CANDIDATES if c != name_column]
    for candidate in candidates:
        try:
            return find_column(ws, candidate)
        except ValueError:
            continue
    raise ValueError(f"None of {candidates!r} found in header row")


def _read_getty_ids(path, sheet_name, name_column, seconds_column, min_seconds):
    """Returns one (clip_id, seconds) tuple per unique qualifying clip."""
    wb = load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    name_col_idx = _find_name_column(ws, name_column)
    seconds_col_idx = find_column(ws, seconds_column)

    seen = set()
    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=name_col_idx).value
        if name is None or str(name).strip() == "":
            continue
        seconds = ws.cell(row=r, column=seconds_col_idx).value
        if not isinstance(seconds, (int, float)) or seconds < min_seconds:
            continue
        clip_id = extract_getty_id(name)
        if clip_id in seen:
            continue
        seen.add(clip_id)
        rows.append((clip_id, seconds))
    return rows


_TITLE_FONT = Font(name="Lato", size=12, color="FF7030A0")
_HEADER_FONT = Font(name="Lato", size=8)
_FORMULA_FONT = Font(name="Lato", size=10)
_CENTER = Alignment(horizontal="center")
_CENTER_WRAP = Alignment(horizontal="center", wrap_text=True)
_DASHED_BOTTOM = Border(bottom=Side(style="dashed"))
_DOTTED_BOTTOM = Border(bottom=Side(style="dotted"))

_FORM_TITLE_FONT = Font(name="Lato", size=20, bold=True, color="FF7030A0")
_FIELD_LABEL_FONT = Font(name="Lato", size=11)
_FIELD_VALUE_FONT = Font(name="Calibri", size=11)
_FIELD_VALUE_ALIGN = Alignment(horizontal="center")

_ASSET_COL_WIDTH = 20
_DURATION_COL_WIDTH = 10

_FORM_TITLE_ROW = 1
_COMPANY_ROW = 2
_PROJECT_ROW = 3
_BROADCASTER_ROW = 4
_RIGHTS_ROW = 5

_BLOCK_TITLE_ROW = 7
_BLOCK_SUBTITLE_ROW = 8
_BLOCK_HEADER_ROW = 9
_BLOCK_FORMULA_ROW = 10
_BLOCK_FIRST_DATA_ROW = 11

_INTRA_BLOCK_GAP = 1    # blank columns between a Video block and its Stills block
_INTER_EPISODE_GAP = 2  # blank columns between one episode's Stills block and the next episode's Video block


def _write_header(ws, production_company, project_name, broadcaster, rights):
    title_cell = ws.cell(row=_FORM_TITLE_ROW, column=1, value="Customer Declaration Form")
    title_cell.font = _FORM_TITLE_FONT

    fields = [
        (_COMPANY_ROW, "Production Company:", production_company),
        (_PROJECT_ROW, "Project Name:", project_name),
        (_BROADCASTER_ROW, "Broadcaster:", broadcaster),
        (_RIGHTS_ROW, "Rights Requested:", rights),
    ]
    for row, label, value in fields:
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.font = _FIELD_LABEL_FONT
        value_cell = ws.cell(row=row, column=2, value=value)
        value_cell.font = _FIELD_VALUE_FONT
        value_cell.alignment = _FIELD_VALUE_ALIGN


def _write_block(ws, col, title, subtitle, rows):
    """Writes one 2-column Asset ID/Duration block starting at `col`.
    `rows` is a list of (clip_id, seconds) tuples, same shape
    _read_getty_ids() returns."""
    asset_col, duration_col = col, col + 1
    asset_letter = get_column_letter(asset_col)
    duration_letter = get_column_letter(duration_col)

    title_cell = ws.cell(row=_BLOCK_TITLE_ROW, column=asset_col, value=title)
    title_cell.font = _TITLE_FONT
    title_cell.alignment = _CENTER
    ws.merge_cells(start_row=_BLOCK_TITLE_ROW, start_column=asset_col, end_row=_BLOCK_TITLE_ROW, end_column=duration_col)

    subtitle_cell = ws.cell(row=_BLOCK_SUBTITLE_ROW, column=asset_col, value=subtitle)
    subtitle_cell.font = _TITLE_FONT
    subtitle_cell.alignment = _CENTER
    ws.merge_cells(start_row=_BLOCK_SUBTITLE_ROW, start_column=asset_col, end_row=_BLOCK_SUBTITLE_ROW, end_column=duration_col)

    for c, label in [(asset_col, "Asset ID"), (duration_col, "Duration")]:
        cell = ws.cell(row=_BLOCK_HEADER_ROW, column=c, value=label)
        cell.font = _HEADER_FONT
        cell.alignment = _CENTER
        cell.border = _DASHED_BOTTOM

    last_data_row = _BLOCK_FIRST_DATA_ROW + len(rows) - 1 if rows else _BLOCK_FIRST_DATA_ROW
    count_cell = ws.cell(
        row=_BLOCK_FORMULA_ROW, column=asset_col,
        value=f"=COUNTA({asset_letter}{_BLOCK_FIRST_DATA_ROW}:{asset_letter}{last_data_row})",
    )
    sum_cell = ws.cell(
        row=_BLOCK_FORMULA_ROW, column=duration_col,
        value=f"=SUM({duration_letter}{_BLOCK_FIRST_DATA_ROW}:{duration_letter}{last_data_row})",
    )
    for cell in (count_cell, sum_cell):
        cell.font = _FORMULA_FONT
        cell.alignment = _CENTER
        cell.border = _DASHED_BOTTOM

    for i, (clip_id, seconds) in enumerate(rows):
        r = _BLOCK_FIRST_DATA_ROW + i
        id_cell = ws.cell(row=r, column=asset_col, value=clip_id)
        id_cell.alignment = _CENTER_WRAP
        ws.cell(row=r, column=duration_col, value=seconds)

    ws.column_dimensions[asset_letter].width = _ASSET_COL_WIDTH
    ws.column_dimensions[duration_letter].width = _DURATION_COL_WIDTH


def build_getty_id_report(
    input_path,
    out_path,
    project_name,
    sheet_name="Getty videos",
    stills_sheet_name="Getty pics",
    name_column="Clip Name",
    seconds_column="Seconds",
    min_seconds=5,
    production_company="KM Record a.s./Big Media",
    broadcaster="",
    rights="in perpetuity/worldwide/all media",
):
    out_name = Path(out_path).name
    files = [f for f in iter_xlsx_files(input_path) if f.name != out_name]

    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    ws_out = wb_out.create_sheet(title="Getty IDs")

    _write_header(ws_out, production_company, project_name, broadcaster, rights)

    counts = {}
    col = 1
    last_used_col = 0
    for file_path in files:
        video_rows = _read_getty_ids(file_path, sheet_name, name_column, seconds_column, min_seconds)
        stills_rows = _read_getty_ids(file_path, stills_sheet_name, name_column, seconds_column, min_seconds)
        counts[file_path.name] = {"video": len(video_rows), "stills": len(stills_rows)}

        title = clean_episode_title(file_path.name)
        _write_block(ws_out, col, title, "Getty Images Video", video_rows)
        stills_col = col + 2 + _INTRA_BLOCK_GAP
        _write_block(ws_out, stills_col, title, "Getty Images Stills", stills_rows)

        last_used_col = stills_col + 1
        col = last_used_col + 1 + _INTER_EPISODE_GAP

    ws_out.freeze_panes = f"A{_BLOCK_FIRST_DATA_ROW}"
    wb_out.save(out_path)
    return counts
