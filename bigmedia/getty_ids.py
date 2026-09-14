"""
Extract Getty clip ids from clip filenames and consolidate them from
several already-sorted workbooks into one simplified Getty IDs report file.

Writes a header with just a Project Name field. For each input file, reads
both the "Getty Videos" and "Getty Stills" sheets (matched case-insensitively;
the stills sheet also falls back to the known alias "Getty pics" if
"Getty Stills"/the caller's --stills-sheet isn't found, since real delivery
files use both names even within the same batch — pass --stills-sheet for any
other custom name), extracts each clip's Getty id from its filename (see
extract_getty_id() for the exact rules), and writes one 2-column block
(Asset ID / Duration) per file per clip type. Episodes appear left to right
in input order, each block followed by blank spacing before the next episode.

Videos and stills are written to separate sheets ("Getty Images Video" /
"Getty Images Stills"). A sheet is only created if it has at least one clip
across all input files — e.g. a project with no stills at all gets no
"Getty Images Stills" sheet.

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
# "(N)" copy marker (with or without a leading space), before or in place
# of the extension -- e.g. "...-640_ADPP (3).MP4".
_COPY_MARKER_RE = re.compile(r'\s*\(\d+\)\s*$')
# The fixed "-640_ADPP" resolution/tool block, plus an optional "(N)" copy
# marker glued to it -- hyphen-joined to the id, so it can't be caught by
# the "_"-separated junk-part walk below.
_ADPP_SUFFIX_RE = re.compile(r'-640_ADPP(?:\s*\(\d+\))?', re.IGNORECASE)

# Known junk tags export/post-process tools bake onto the id as a trailing
# "_word" chain -- explicitly whitelisted (mirrors dedupe.py's codec/
# post-process suffix regexes) rather than "any non-numeric segment is
# junk", because real ids also end in non-numeric-looking segments that are
# NOT junk (e.g. "1B010728_t010" -- "t010" is part of the id, not a tag).
_JUNK_PART_RE = re.compile(
    r'^(?:Apple|ProRes|APPLEPRORES(?:HQ|LT)?|422|444|4444|HQ|LT|Proxy|DNxHD|DNxHR|'
    r'H\.?264|H\.?265|HEVC|AVC|XDCAM|MPEG-?[24]?|Denoise|Deflicker|NTSC|'
    r'AvidRetime(?:-\d+)?|S\d+|upscale\d*|AIUPSCALE\d*|SM\d*|DFR\d*)$',
    re.IGNORECASE,
)


def extract_getty_id(name: str) -> str:
    """
    Pull the Getty id out of a clip filename. Rules, derived from real
    examples (see tests/test_getty_ids.py for the full regression list):
      - strip a leading "GettyImages-"/"GETTYIMAGES-" prefix
      - strip the fixed "-640_ADPP" block (and a "(N)" copy marker glued to
        it) wherever it appears, e.g. "...-640_ADPP (3).MP4" -> "..."
      - a leading "mr_"/"MR_" tag right after that prefix is part of the id,
        not metadata — keep it as-is, e.g. "mr_00108323.mov" -> "mr_00108323"
      - if a .mov/.mp4/.jpg/.new extension appears anywhere, cut the string
        at the start of that extension, dropping it and everything after it
        (including trailing " 25"-style suffixes, or a re-render marker
        like ".new.02"); then drop a trailing " (N)" copy marker.
      - walk the "_"-separated parts after the first and drop a part plus
        everything after it the moment one matches a known junk tag
        (_JUNK_PART_RE, e.g. "_Apple_ProRes_422", "_APPLEPRORESHQ",
        "_Denoise_02", "_S000_upscale01", "_AIUPSCALE", "_SM01",
        "_DFR01_OK"). Purely numeric parts (e.g. "_0003") and any other
        non-junk part (e.g. "_t010") are kept as part of the id.
      - if no extension is present, also drop a bare trailing hyphen with
        nothing after it (e.g. "96410456-" -> "96410456"); a trailing
        " 2"-style suffix is kept.
    """
    s = (name or "").strip()
    s = _GETTY_PREFIX_RE.sub('', s)
    s = _ADPP_SUFFIX_RE.sub('', s)

    m = _EXT_RE.search(s)
    if m:
        s = s[:m.start()]
    s = _COPY_MARKER_RE.sub('', s)

    parts = s.split('_')
    core = [parts[0]]
    for part in parts[1:]:
        if _JUNK_PART_RE.match(part):
            break
        core.append(part)
    s = '_'.join(core)

    if not m:
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


def getty_ids_filename(project_name: str) -> str:
    """The report's default filename, e.g. "Wild Return" -> "Wild Return - Getty_IDs.xlsx"."""
    return f"{project_name} - Getty_IDs.xlsx"


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

_FORM_TITLE_FONT = Font(name="Lato", size=20, bold=True, color="FF7030A0")
_FIELD_LABEL_FONT = Font(name="Lato", size=11)
_FIELD_VALUE_FONT = Font(name="Calibri", size=11)
_FIELD_VALUE_ALIGN = Alignment(horizontal="center")

_ASSET_COL_WIDTH = 20
_DURATION_COL_WIDTH = 10

_FORM_TITLE_ROW = 1
_PROJECT_ROW = 2

_BLOCK_TITLE_ROW = 4
_BLOCK_SUBTITLE_ROW = 5
_BLOCK_HEADER_ROW = 6
_BLOCK_FORMULA_ROW = 7
_BLOCK_FIRST_DATA_ROW = 8

_INTER_EPISODE_GAP = 2  # blank columns between one episode's block and the next

_VIDEO_SHEET_NAME = "Getty Images Video"
_STILLS_SHEET_NAME = "Getty Images Stills"


def _write_header(ws, project_name):
    title_cell = ws.cell(row=_FORM_TITLE_ROW, column=1, value="Customer Declaration Form")
    title_cell.font = _FORM_TITLE_FONT

    label_cell = ws.cell(row=_PROJECT_ROW, column=1, value="Project Name:")
    label_cell.font = _FIELD_LABEL_FONT
    value_cell = ws.cell(row=_PROJECT_ROW, column=2, value=project_name)
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


def _write_sheet(wb_out, sheet_name, project_name, subtitle, per_file_rows):
    """Creates one sheet and writes one block per (title, rows) entry in
    per_file_rows, left to right in input order."""
    ws = wb_out.create_sheet(title=sheet_name)
    _write_header(ws, project_name)

    col = 1
    for title, rows in per_file_rows:
        _write_block(ws, col, title, subtitle, rows)
        col += 2 + _INTER_EPISODE_GAP

    ws.freeze_panes = f"A{_BLOCK_FIRST_DATA_ROW}"
    return ws


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
    include_video=True,
    include_stills=True,
):
    out_name = Path(out_path).name
    files = [f for f in iter_xlsx_files(input_path) if f.name != out_name]

    counts = {}
    per_file_video = []
    per_file_stills = []
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
        per_file_video.append((title, video_rows))
        per_file_stills.append((title, stills_rows))

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    # Only write a sheet for a clip type if it actually has at least one
    # clip somewhere -- e.g. a project with no stills gets no Stills sheet.
    if include_video and any(rows for _, rows in per_file_video):
        _write_sheet(wb_out, _VIDEO_SHEET_NAME, project_name, _VIDEO_SHEET_NAME, per_file_video)
    if include_stills and any(rows for _, rows in per_file_stills):
        _write_sheet(wb_out, _STILLS_SHEET_NAME, project_name, _STILLS_SHEET_NAME, per_file_stills)

    if not wb_out.sheetnames:
        fallback_name = _VIDEO_SHEET_NAME if include_video else _STILLS_SHEET_NAME
        _write_sheet(wb_out, fallback_name, project_name, fallback_name, per_file_video if include_video else per_file_stills)

    wb_out.save(out_path)
    return counts
