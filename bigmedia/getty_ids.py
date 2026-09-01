"""
Extract Getty clip ids from clip filenames and consolidate them from
several already-sorted workbooks into one Customer Declaration Form-styled
report file. The form matches Getty's template so it can be pasted/sent back
to Getty with no manual reformatting.

Writes a header section with Production Company, Project Name, Broadcaster,
and Rights fields. For each input file, reads both the "Getty Videos" and
"Getty Stills" sheets (matched case-insensitively; the stills sheet also
falls back to the known alias "Getty pics" if "Getty Stills"/the caller's
--stills-sheet isn't found, since real delivery files use both names even
within the same batch — pass --stills-sheet for any other custom name), extracts
each clip's Getty id from its filename (see extract_getty_id() for the
exact rules), and writes two
adjacent 2-column blocks (Asset ID / Duration) per file: one for Video clips,
one for Stills. Episodes appear left to right in input order, each pair of
blocks followed by blank spacing before the next episode.

Only UNIQUE video clips with a total duration of 5+ seconds (--min-seconds,
--max-seconds) are included. Real delivery files (and our own `group`
command's output) only populate the "Seconds" column once per clip — on
the last row of a duplicate run, blank on the earlier rows of that same
run — so filtering to "Seconds present and >= min_seconds" naturally
selects one row per unique clip. An explicit id-based dedup is layered on
top as a safety net in case a file doesn't follow that convention.

Stills have no meaningful "duration" -- min/max-seconds never filters
them, every uniquely-named still is included regardless of what (if
anything) is in its Seconds cell.

--videos-only / --stills-only restrict extraction to just one sheet
(mutually exclusive); by default both are processed.

After the last episode's blocks, fixed gaps place two text boxes side-by-side:
an "Instructions:" box explaining the form's usage, and an "Important Notes
on Licensing:" box with Getty's licensing terms.
"""
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .xlsx_utils import find_column, iter_xlsx_files

_GETTY_PREFIX_RE = re.compile(r'^GettyImages-', re.I)
_EXT_RE = re.compile(r'\.(mov|mp4|jpg|new)\b', re.I)
_KOPIE_SUFFIX_RE = re.compile(r'\s*\(kopie\)\s*$', re.I)

# Known junk tags export/post-process tools bake onto the id as a trailing
# "_word" chain -- explicitly whitelisted (mirrors dedupe.py's codec/
# post-process suffix regexes) rather than "any non-numeric segment is
# junk", because real ids also end in non-numeric-looking segments that are
# NOT junk (e.g. "1B010728_t010" -- "t010" is part of the id, not a tag).
_JUNK_PART_RE = re.compile(
    r'^(?:Apple|ProRes|422|444|4444|HQ|LT|Proxy|DNxHD|DNxHR|H\.?264|H\.?265|HEVC|AVC|'
    r'XDCAM|MPEG-?[24]?|Denoise|Deflicker|NTSC|AvidRetime(?:-\d+)?|S\d+|upscale\d*)$',
    re.IGNORECASE,
)


def extract_getty_id(name: str) -> str:
    """
    Pull the Getty id out of a clip filename. Rules, derived from real
    examples (see tests/test_getty_ids.py for the full regression list):
      - strip a leading "GettyImages-"/"GETTYIMAGES-" prefix
      - a leading "mr_"/"MR_" tag right after that prefix is part of the id,
        not metadata — keep it as-is, e.g. "mr_00108323.mov" -> "mr_00108323"
      - if a .mov/.mp4/.jpg/.new extension appears anywhere, cut the string
        at the start of that extension, dropping it and everything after it
        (including trailing " 25"-style suffixes, or a re-render marker
        like ".new.02"). Within what's left, walk the "_"-separated parts
        after the first and drop a part plus everything after it the moment
        one matches a known junk tag (_JUNK_PART_RE, e.g. "_Apple_ProRes_422",
        "_Denoise_02", "_S000_upscale01"). Purely numeric parts (e.g.
        "_0003") and any other non-junk part (e.g. "_t010") are kept as
        part of the id.
      - if no extension is present, keep the remainder as-is (including any
        trailing " 2"-style suffix), except a bare trailing hyphen with
        nothing after it (e.g. "96410456-" -> "96410456").
    """
    s = (name or "").strip()
    s = _GETTY_PREFIX_RE.sub('', s)

    m = _EXT_RE.search(s)
    if m:
        s = s[:m.start()]
        parts = s.split('_')
        core = [parts[0]]
        for part in parts[1:]:
            if _JUNK_PART_RE.match(part):
                break
            core.append(part)
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


def _find_sheet(wb, sheet_name):
    """Sheet titles vary in case across real delivery/sort output (e.g.
    "Getty Videos" vs "Getty videos") — match case-insensitively, same
    convention as group_duplicates.py."""
    for title in wb.sheetnames:
        if title.strip().lower() == sheet_name.strip().lower():
            return title
    return None


# Real delivery files sometimes use an entirely different word for the
# stills sheet, not just different casing (e.g. "Getty pics" instead of
# "Getty Stills") — a plain case-insensitive match on the caller's chosen
# name won't catch that, so fall back through known aliases same as
# NAME_COLUMN_ALIASES above.
_STILLS_SHEET_ALIASES = ("Getty Stills", "Getty pics")


def _find_sheet_any(wb, sheet_name, aliases=()):
    candidates = [sheet_name] + [a for a in aliases if a.strip().lower() != sheet_name.strip().lower()]
    for candidate in candidates:
        title = _find_sheet(wb, candidate)
        if title:
            return title
    return None


def _read_getty_ids(path, sheet_name, name_column, seconds_column, min_seconds, max_seconds=None,
                     sheet_aliases=(), filter_by_seconds=True):
    """Returns one (clip_id, seconds) tuple per unique qualifying clip.

    Stills have no meaningful duration, so callers pass filter_by_seconds=
    False for them: every uniquely-named row is included regardless of its
    Seconds cell (even if that column is missing or blank), min/max-seconds
    is a video-only concept."""
    wb = load_workbook(path, data_only=True)
    actual_sheet_name = _find_sheet_any(wb, sheet_name, sheet_aliases)
    if actual_sheet_name is None:
        return []
    ws = wb[actual_sheet_name]
    name_col_idx = _find_name_column(ws, name_column)
    try:
        seconds_col_idx = find_column(ws, seconds_column)
    except ValueError:
        seconds_col_idx = None

    seen = set()
    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=name_col_idx).value
        if name is None or str(name).strip() == "":
            continue
        seconds = ws.cell(row=r, column=seconds_col_idx).value if seconds_col_idx else None
        if filter_by_seconds:
            if not isinstance(seconds, (int, float)) or seconds < min_seconds:
                continue
            if max_seconds is not None and seconds >= max_seconds:
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

_BOX_GAP = 2       # blank columns between the last episode block and the Instructions box
_BOX_INNER_GAP = 1  # blank columns between the Instructions box and the Notes box
_INSTRUCTIONS_WIDTH = 6
_NOTES_WIDTH = 8

_BOX_ROW = 5
_BOX_HEADER_ROWSPAN = 2
_BOX_BODY_ROWSPAN = 26

_BOX_HEADER_FONT = Font(name="Lato", size=15, bold=True, color="FF7030A0")
_BOX_HEADER_ALIGN = Alignment(horizontal="left", vertical="center")
_BOX_BODY_FONT = Font(name="Lato", size=11)
_BOX_BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)

_INSTRUCTIONS_TEXT = (
    "Please divide content into appropriate asset type.\n\n"
    "Under asset ID, please enter the ID listed on the Getty Images website "
    "for this item - this will either be a ‘Creative #’, ‘Editorial #’ for "
    "stills or ‘Clip #’ for online video items. For offline items, your clip "
    "ID should be entered here. \n\n"
    "For your video items, under 'duration' please enter the number of "
    "seconds used of this video within your final edit.\n\n"
    "More content subcategories are in hidden columns, please only expand "
    "if you need to declare BBC Sport content, NBC Premium or Standard OR "
    "alternative Getty Images options.  "
)

_NOTES_TEXT = (
    "Upon receipt of the Proposed Usage Declaration form, Getty Images will "
    "check availability of all itemised content and shall inform Customer "
    "as soon as reasonably practicable whether content is available for "
    "license, i.e. after checking that content is still represented and "
    "available for licensing by Getty Images, product specialist team will "
    "update once confirmed. Availability of content is not guaranteed, "
    "Customer shall not finalise the Production Title until Getty Images "
    "has confirmed availability of licensing rights. Content shall be "
    "deemed licensed upon Getty Images confirming it is available for "
    "license in response to receiving a Proposed Usage Declaration.       \n"
    "                                                                                                                                     "
    "For all offline content, once the master material is supplied, the applicable\n"
    "license fee and all technical charges are payable regardless of "
    "whether the master material is used or not.\n"
    "In addition, further approval and delivery mechanisms apply for all "
    "offline BBC Motion Gallery, BBC Sport and NBC video collections, as "
    "outlined in agreement contract."
)


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


def _write_text_box(ws, col, width, header_text, body_text, header_border):
    end_col = col + width - 1
    header_row_end = _BOX_ROW + _BOX_HEADER_ROWSPAN - 1
    header_cell = ws.cell(row=_BOX_ROW, column=col, value=header_text)
    header_cell.font = _BOX_HEADER_FONT
    header_cell.alignment = _BOX_HEADER_ALIGN
    header_cell.border = header_border
    ws.merge_cells(start_row=_BOX_ROW, start_column=col, end_row=header_row_end, end_column=end_col)

    body_row = header_row_end + 1
    body_row_end = body_row + _BOX_BODY_ROWSPAN - 1
    body_cell = ws.cell(row=body_row, column=col, value=body_text)
    body_cell.font = _BOX_BODY_FONT
    body_cell.alignment = _BOX_BODY_ALIGN
    ws.merge_cells(start_row=body_row, start_column=col, end_row=body_row_end, end_column=end_col)


def build_getty_id_report(
    input_path,
    out_path,
    project_name,
    sheet_name="Getty Videos",
    stills_sheet_name="Getty Stills",
    name_column="Clip Name",
    seconds_column="Seconds",
    min_seconds=5,
    max_seconds=None,
    production_company="KM Record a.s./Big Media",
    broadcaster="",
    rights="in perpetuity/worldwide/all media",
    include_video=True,
    include_stills=True,
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
        video_rows = (
            _read_getty_ids(file_path, sheet_name, name_column, seconds_column, min_seconds, max_seconds)
            if include_video else []
        )
        stills_rows = (
            _read_getty_ids(
                file_path, stills_sheet_name, name_column, seconds_column, min_seconds, max_seconds,
                sheet_aliases=_STILLS_SHEET_ALIASES, filter_by_seconds=False,
            )
            if include_stills else []
        )
        counts[file_path.name] = {"video": len(video_rows), "stills": len(stills_rows)}

        title = clean_episode_title(file_path.name)
        blocks = []
        if include_video:
            blocks.append(("Getty Images Video", video_rows))
        if include_stills:
            blocks.append(("Getty Images Stills", stills_rows))

        block_col = col
        for i, (subtitle, rows) in enumerate(blocks):
            if i > 0:
                block_col += _INTRA_BLOCK_GAP
            _write_block(ws_out, block_col, title, subtitle, rows)
            block_col += 2

        last_used_col = block_col - 1
        col = last_used_col + 1 + _INTER_EPISODE_GAP

    box_start_col = last_used_col + 1 + _BOX_GAP
    _write_text_box(ws_out, box_start_col, _INSTRUCTIONS_WIDTH, "Instructions:", _INSTRUCTIONS_TEXT, _DASHED_BOTTOM)
    notes_start_col = box_start_col + _INSTRUCTIONS_WIDTH + _BOX_INNER_GAP
    _write_text_box(ws_out, notes_start_col, _NOTES_WIDTH, "Important Notes on Licensing:", _NOTES_TEXT, _DOTTED_BOTTOM)

    ws_out.freeze_panes = f"A{_BLOCK_FIRST_DATA_ROW}"
    wb_out.save(out_path)
    return counts
