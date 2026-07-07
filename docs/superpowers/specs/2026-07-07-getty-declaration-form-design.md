# Getty Customer Declaration Form output — design

## Context

`build_getty_id_report()` currently reads one sheet ("Getty videos") from
each sorted input workbook and writes one 2-column (Asset ID / Duration)
block per file into a plain "Getty IDs" sheet. Valentina attached a real
Getty "Customer Declaration Form" (`Customer Declaration Form.xlsx`) and
wants the generated report to look like that form — header fields,
Instructions/Important Notes boxes, and Video + Stills blocks per episode
— so it can be sent to Getty without manual reformatting.

Inspecting the attached file confirmed the current block styling (fonts,
sizes, purple `FF7030A0` title color, dashed borders under header/formula
rows) already matches the real form closely — it was clearly derived from
a form like this originally. The gap is: no header section, no
Instructions/Notes boxes, and only one asset-type block (Video) instead of
two (Video + Stills) per episode.

## Decisions (confirmed with Valentina)

- **No binary template bundled.** Header fields, Instructions text, and
  Important Notes text become Python string/style constants in
  `getty_ids.py`, matching the pattern already used for the block styling.
  Avoids committing a client's real Production Company/Project data baked
  into a binary asset, and keeps everything in one editable, testable file.
- **Header fields are CLI-configurable:**
  - `--production-company` (default `"KM Record a.s./Big Media"`)
  - `--project-name` (**required** — changes every job, a blank value
    would look broken on a real form)
  - `--broadcaster` (optional, blank default — matches the example, where
    this row was empty)
  - `--rights` (default `"in perpetuity/worldwide/all media"`)
- **Both Getty sheets are read per input file:** `"Getty videos"` (video
  ids) and `"Getty pics"` (still ids), same extraction/filter rules
  (`_read_getty_ids`, `min_seconds`) for both. Each becomes its own 2-column
  block (Asset ID / Duration), Video block immediately left of its
  matching Stills block, same episode title over each (independently, not
  merged across both — confirmed from the real file: "EP1 - Eiffel Tower"
  is typed separately into each block's own title cell, not one merged
  title spanning both blocks).
- **Formula ranges are always correct.** The real file's 2nd episode block
  has a copy-paste bug (`N14` sums `N14:N3602`, including itself, instead
  of `N15:N3602` like the 1st block). Generated blocks always use the
  correct data range — the bug is not replicated.
- **Instructions/Notes boxes are positioned dynamically**, a fixed gap
  after the last episode's Stills block, so the layout is correct for any
  number of input files (not just the 2 in the example).

## Layout

Row layout (previously: block title started at row 1). Now:

```
Row 1   "Customer Declaration Form"           (title, big, bold, purple)
Row 2   Production Company:  <value>
Row 3   Project Name:        <value>
Row 4   Broadcaster:         <value>
Row 5   Rights Requested:    <value>
Row 6   (blank spacer)
Row 7   <episode title>          <episode title>       ...one pair of cells
Row 8   Getty Images Video       Getty Images Stills    per 2-col block
Row 9   Asset ID | Duration      Asset ID | Duration
Row 10  =COUNTA  | =SUM          =COUNTA  | =SUM
Row 11+ <id>     | <seconds>     <id>     | <seconds>
```

Column layout, left to right, repeating per input file:

```
[Video: Asset ID, Duration] [1-col gap] [Stills: Asset ID, Duration] [2-col gap] ... next file
```

After the last file's Stills block: 2-col gap, then the Instructions box
(6 cols wide), 1-col gap, then the Important Notes box (8 cols wide) —
both anchored at row 5 (matching the example's independent vertical
anchor above the block titles), header cell 2 rows tall, body cell merged
26 rows tall, `wrap_text=True`, verbatim legal text copied from the
example file.

Column widths: Asset ID 20, Duration 10 (existing constants, applied
uniformly — the example's per-position widths were inconsistent/manually
resized and aren't meaningful to replicate).

## Function signature

```python
def build_getty_id_report(
    input_path,
    out_path,
    project_name,                                    # required, no default
    sheet_name="Getty videos",
    stills_sheet_name="Getty pics",
    name_column="Clip Name",
    seconds_column="Seconds",
    min_seconds=5,
    production_company="KM Record a.s./Big Media",
    broadcaster="",
    rights="in perpetuity/worldwide/all media",
):
```

Returns `{filename: {"video": n, "stills": n}}` — a breaking change from
today's flat `{filename: n}`, needed so the CLI summary can report the
video/stills split per file (e.g. "12 clips (9 video, 3 stills)").

## Testing

`tests/test_getty_ids.py` gets rewritten for the new row/column offsets
and the mandatory `project_name` arg — every existing assertion's row
numbers shift down by 6, and column numbers for a "2nd file" shift right
to account for the Stills block now sitting between one file's blocks and
the next file's. New cases added for: stills-only sheet present but video
sheet missing, header field values landing in the right cells, Instructions
/Important Notes box text and position (including position shifting
correctly with 1 vs 3 input files), and the corrected (non-buggy) SUM
range.
