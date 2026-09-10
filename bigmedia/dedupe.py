"""
Duplicate detection for clip names.

dedup_key() is the original, proven logic (strips extension and trailing
"(1)"-style copy markers so "Clip_A.mov" and "Clip_A (1).mxf" collapse to
the same key). It also swallows a trailing " <digits>" annotation right
after the extension (e.g. a stray frame-rate/version tag), so
"GettyImages-627-122.mov" and "GettyImages-627-122.mov 25" collapse to the
same key too. It also strips a trailing known codec/export tag right before the
extension -- some export tools bake the codec into the filename on
re-export of the same clip. Both the spaced-out form ("_Apple_ProRes_422")
and the run-together form ("_APPLEPRORESHQ") match. Same for a Getty
upscale/segment tag ("_S000_upscale01", bare "_upscale01", "_AIUPSCALE"),
a post-process tool tag ("_AvidRetime-10383301", "_Denoise", "_Denoise_chr2",
"_Deflicker", "_NTSC"), a submaster tag ("_SM"/"_SMnn"), a "_DFR"/"_DFRnn"
tag with an optional "_OK" QC-pass marker, and the fixed "-640_ADPP"
resolution/tool block.

It also strips Getty QC/review-export junk: "Render" and "Render <N>" take
markers (space- or underscore-joined), "_prob4" problem-pass tags, and
"_comp"/"_comp_<N>" composite-render tags -- these pipelines re-append the
extension on every pass (e.g. "....mov Render 1_prob4.mov" or "....mov
Render.mov Render_1_prob4.mov").

Every one of these strippers, plus the extension and the " (1)"-style copy
marker, runs in a single loop until the name stops shrinking -- so any
combination of tags collapses regardless of the order they were appended
in (e.g. "..._SM01_APPLEPRORESHQ.MOV", "...-640_ADPP (3).MP4").

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

from .xlsx_utils import find_column, find_column_any, capture_header_template, write_header_row, write_data_row, copy_sheet_verbatim

# Real delivery files aren't consistent about the clip-name header: some
# episodes use "Name", others "Clip Name" for the same data. Whichever one
# is passed as name_column is tried first; this covers the other case.
NAME_COLUMN_ALIASES = ("Clip Name", "Name")

# Known codec/export tags some tools bake into the filename right before the
# extension when re-exporting the same clip. Whitelisted deliberately --
# stripping *any* trailing "_word" segment would risk collapsing two
# genuinely different clips whose ids merely contain an underscore (e.g.
# "GettyImages-356-30_1234.mov"). The separators between the Apple / ProRes /
# quality words are optional so both the spaced-out form ("_Apple_ProRes_422HQ")
# and the run-together form ("_APPLEPRORESHQ") match.
_CODEC_SUFFIX_RE = re.compile(
    r'_(?:Apple[_ ]?)?'
    r'(?:ProRes[_ ]?(?:(?:422|4444)(?:[_ ]?(?:HQ|LT))?|Proxy|HQ|LT)?'
    r'|DNxHD|DNxHR|H\.?264|H\.?265|HEVC|AVC|XDCAM|MPEG-?[24]?)$',
    re.IGNORECASE,
)

# Getty re-exports/upscales sometimes bake an optional segment marker
# ("_S000") and an "_upscaleNN" / "_AIUPSCALE" tag onto the same source id,
# e.g. "GettyImages-650878972_S000_upscale01.mov" -- same clip as
# "GettyImages-650878972_upscale01.mov" and "GettyImages-650878972.mov".
_UPSCALE_SUFFIX_RE = re.compile(r'(?:_S\d+)?_(?:AI[_ ]?)?upscale\d*$', re.IGNORECASE)

# Newer export/post-process tags baked onto the same source clip on
# re-render. Each is only ever a trailing "_"-segment, so anchoring to the
# end is safe: "_SM"/"_SMnn" (submaster), "_DFR"/"_DFRnn" with an optional
# "_OK" QC-pass marker, and the fixed "-640_ADPP" resolution/tool block.
_SUBMASTER_SUFFIX_RE = re.compile(r'_SM\d*$', re.IGNORECASE)
_DFR_SUFFIX_RE = re.compile(r'_DFR\d*(?:_OK)?$', re.IGNORECASE)
_ADPP_SUFFIX_RE = re.compile(r'-640_ADPP$', re.IGNORECASE)

# Post-process tools bake their own tag onto the same source id when a clip
# is re-exported through them -- not a different clip, e.g.
# "GettyImages-551409951_AvidRetime-10383301.mov" is the same clip as
# "GettyImages-551409951.mov". "_Denoise" may itself carry a trailing
# "_chr<digits>" channel/pass marker (e.g. "_Denoise_chr2").
_POST_PROCESS_SUFFIX_RE = re.compile(
    r'_(?:AvidRetime-\d+|Denoise(?:_chr\d+)?|Deflicker|NTSC)$',
    re.IGNORECASE,
)

# Extension + optional trailing " 25"-style tag, same pattern used below.
_EXT_RE = re.compile(r'\.[A-Za-z0-9]{2,4}(?:\s+\d+)?$')

# " (1)"-style copy marker (with or without a leading space).
_COPY_MARKER_RE = re.compile(r'\s*\(\d+\)\s*$')

# Getty QC/review-export junk tags -- see module docstring. Each pipeline
# pass re-appends its own extension, so these are stripped in a loop
# together with _EXT_RE rather than a single pass.
_RENDER_SUFFIX_RE = re.compile(r'[ _]Render(?:[ _]\d+)?$', re.IGNORECASE)
_PROB4_SUFFIX_RE = re.compile(r'_prob4$', re.IGNORECASE)
_COMP_SUFFIX_RE = re.compile(r'_comp(?:_\d+)?$', re.IGNORECASE)

# Every trailing-noise stripper, applied in a loop until the name stops
# shrinking. Looping (rather than one pass each) is what lets any
# combination of tags collapse regardless of the order they were appended
# in -- e.g. "..._SM01_APPLEPRORESHQ.MOV" or "...-640_ADPP_AIUPSCALE (2).MP4".
_NOISE_SUFFIX_RES = (
    _EXT_RE,
    _COPY_MARKER_RE,
    _RENDER_SUFFIX_RE,
    _PROB4_SUFFIX_RE,
    _COMP_SUFFIX_RE,
    _CODEC_SUFFIX_RE,
    _UPSCALE_SUFFIX_RE,
    _POST_PROCESS_SUFFIX_RE,
    _SUBMASTER_SUFFIX_RE,
    _DFR_SUFFIX_RE,
    _ADPP_SUFFIX_RE,
)


def dedup_key(name: str) -> str:
    base = (name or "").strip()
    while True:
        new_base = base
        for suffix_re in _NOISE_SUFFIX_RES:
            new_base = suffix_re.sub('', new_base)
        new_base = new_base.strip()
        if new_base == base:
            return base
        base = new_base


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

    name_candidates = [name_column] + [a for a in NAME_COLUMN_ALIASES if a != name_column]
    name_col_idx = find_column_any(ws_src, name_candidates)
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
