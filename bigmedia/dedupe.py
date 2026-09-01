"""
Duplicate detection for clip names.

dedup_key() is the original, proven logic (strips extension and trailing
"(1)"-style copy markers so "Clip_A.mov" and "Clip_A (1).mxf" collapse to
the same key). It also swallows a trailing " <digits>" annotation right
after the extension (e.g. a stray frame-rate/version tag), so
"GettyImages-627-122.mov" and "GettyImages-627-122.mov 25" collapse to the
same key too. It also strips a trailing known codec/export tag (e.g.
"_Apple_ProRes_422") right before the extension -- some export tools bake
the codec into the filename on re-export of the same clip. Same for a
Getty upscale/segment tag (e.g. "_S000_upscale01" or bare "_upscale01") and
for a post-process tool tag (e.g. "_AvidRetime-10383301", "_Denoise",
"_Denoise_chr2", "_Deflicker", "_NTSC").

It also strips Getty QC/review-export junk: "Render" and "Render <N>" take
markers (space- or underscore-joined), "_prob4" problem-pass tags, and
"_comp"/"_comp_<N>" composite-render tags -- these pipelines re-append the
extension on every pass (e.g. "....mov Render 1_prob4.mov" or "....mov
Render.mov Render_1_prob4.mov"), so this strip loops together with the
extension strip until nothing more matches, instead of running once.

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
# extension when re-exporting the same clip (e.g. "..._Apple_ProRes_422.mov").
# Whitelisted deliberately -- stripping *any* trailing "_word" segment would
# risk collapsing two genuinely different clips whose ids merely contain an
# underscore (e.g. "GettyImages-356-30_1234.mov").
_CODEC_SUFFIX_RE = re.compile(
    r'_(?:Apple_)?(?:ProRes(?:_(?:422|4444)(?:HQ|LT)?|_Proxy)?|DNxHD|DNxHR|H\.?264|H\.?265|HEVC|AVC|XDCAM|MPEG-?[24]?)$',
    re.IGNORECASE,
)

# Getty re-exports/upscales sometimes bake an optional segment marker
# ("_S000") and an "_upscaleNN" tag onto the same source id, e.g.
# "GettyImages-650878972_S000_upscale01.mov" -- same clip as
# "GettyImages-650878972_upscale01.mov" and "GettyImages-650878972.mov".
_UPSCALE_SUFFIX_RE = re.compile(r'(?:_S\d+)?_upscale\d*$', re.IGNORECASE)

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

# Getty QC/review-export junk tags -- see module docstring. Each pipeline
# pass re-appends its own extension, so these are stripped in a loop
# together with _EXT_RE rather than a single pass.
_RENDER_SUFFIX_RE = re.compile(r'[ _]Render(?:[ _]\d+)?$', re.IGNORECASE)
_PROB4_SUFFIX_RE = re.compile(r'_prob4$', re.IGNORECASE)
_COMP_SUFFIX_RE = re.compile(r'_comp(?:_\d+)?$', re.IGNORECASE)


def _strip_getty_review_export_junk(base):
    while True:
        new_base = _EXT_RE.sub('', base)
        new_base = _RENDER_SUFFIX_RE.sub('', new_base)
        new_base = _PROB4_SUFFIX_RE.sub('', new_base)
        new_base = _COMP_SUFFIX_RE.sub('', new_base)
        if new_base == base:
            return new_base
        base = new_base


def dedup_key(name: str) -> str:
    base = (name or "").strip()
    base = _strip_getty_review_export_junk(base)                 # strip extension(s) + Render/prob4/comp junk
    base = _CODEC_SUFFIX_RE.sub('', base)                        # strip known codec/export tag
    base = _UPSCALE_SUFFIX_RE.sub('', base)                      # strip known upscale/segment tag
    base = _POST_PROCESS_SUFFIX_RE.sub('', base)                 # strip known post-process tool tag
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
