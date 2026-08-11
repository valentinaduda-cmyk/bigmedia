# FU grid output — design

## Context

Valentina hands over an already-sorted workbook (output of `bigmedia
sort`) and needs a "FU grid" — the legal follow-up sheet built from that
workbook's `3rd parties` sheet. Today this is made by hand: dedupe the
clips, count how many times each one is used, sum the durations of all
its uses, rename/reorder a subset of columns, and add empty columns the
legal team fills in later.

The attached example, `data/in/FU grid example.xlsx` (single sheet, "FU
grid", 15 data rows), was made by hand from
`data/in/9_11/260619_SPTBR_11_..._sorted.xlsx`. It is the formatting and
column-order reference, not a byte-for-byte target: some of its cell
values were hand-edited (clip names uppercased, spaces swapped for
underscores inconsistently, some timecodes a frame off from the source
EDL) and are not reproducible by code. Where the example and the source
EDL disagree, the source EDL wins.

## Decisions (confirmed with Valentina)

- **Grouping key is the exact `Clip Name` string.** Not `dedup_key()`.
  This is a deliberate departure from the example file: there, the "Evan
  Fairbanks" row shows `TOTAL USES` 11, which only happens if
  `.mp4` / `_upscale.mov` / `_S000_upscale.mov` variants collapse
  together. With exact matching that group splits into 3 rows (8 + 2 +
  1 uses). Valentina chose exact matching — the grid is a legal
  check-list, so two files that differ at all are checked separately.
- **Single-use clips are kept**, with `TOTAL USES` = 1 and `TOTAL
  DURATION if multiple uses` filled in anyway — for a 1-use clip it
  repeats that clip's own duration. Despite the column name, Valentina
  wants no blank cells in that column. The example contains no 1-use
  rows, so it gives no guidance here.
- **The surviving row of a group is the one with the earliest `Sequence
  In`.** `TC IN`, `TC OUT`, `CLIP DURATION`, `SOURCE DURATION` and
  `CLIP NAME` all come from that one row — never mixed across rows.
- **Grid rows are ordered by that `TC IN` ascending**, matching the
  example.
- **`CLIP NAME` is verbatim from the EDL**, extension included. The
  example's uppercase/underscored names are hand edits and are not
  reproduced.
- **Output sheet order is `3rd parties` first, then `FU grid`** —
  follows the repo convention that every workbook-producing command
  writes its backup copy first.
- **fps defaults to 25**, overridable with `--fps`, same convention as
  `bigmedia group`. Only `Clip Duration` is summed; `Source Duration`
  is copied as text and never parsed (real files carry values like
  `00:00:04:42`, i.e. a different rate — parsing them would be wrong).

## Interface

```python
fu_grid_workbook(
    src_path,
    out_path,
    sheet="3rd parties",
    name_column="Clip Name",
    duration_column="Clip Duration",
    fps=25,
) -> {"rows_in": int, "unique_clips": int, "multi_use_clips": int}
```

New module `bigmedia/fu_grid.py`. CLI wrapper `bigmedia fu-grid IN.xlsx
OUT.xlsx [--sheet] [--name-column] [--duration-column] [--fps]` in
`cli.py`, thin, printing the returned counts so a run is easy to
sanity-check.

Column lookups reuse `xlsx_utils.find_column_any`, so the existing
`Clip Name`/`Name` and `Clip Duration`/`Duration` header aliases keep
working. The source sheet is matched case-insensitively against
`sheet`.

## Output layout

Sheet 1, `3rd parties`: verbatim copy of the source sheet via
`xlsx_utils.copy_sheet_verbatim`.

Sheet 2, `FU grid`: 13 columns, in the example's order.

| # | Header | Source |
|---|--------|--------|
| A | `TC IN` | `Sequence In` of the surviving row |
| B | `TC OUT` | `Sequence Out` of the surviving row |
| C | `CLIP DURATION` | `Clip Duration` of the surviving row |
| D | `TOTAL USES` | group size (int) |
| E | `TOTAL DURATION if multiple uses` | sum of the group's `Clip Duration`, as `HH:MM:SS:FF`; equals column C when size is 1 |
| F | `SOURCE DURATION` | `Source Duration` of the surviving row |
| G | `SCREENSHOTS` | empty |
| H | `URL LINK` | empty |
| I | `PREVIOUS LEGAL CHECK` | empty |
| J | `SOURCE` | empty |
| K | `CLIP NAME` | `Clip Name` of the surviving row, verbatim |
| L | `FINAL FU LEGAL NOTE` | empty |
| M | `PREVIEW: ` | empty (trailing space in the header is verbatim from the example) |

## Styling (cloned from the example)

- Header row: Calibri 11 bold, solid fill `FFCFE2F3`, thin border all
  sides, centered horizontally and vertically, wrap on, row height 43.2.
- Data rows: solid fill `FFF4CCCC` across the row, except column I
  (`PREVIOUS LEGAL CHECK`), which is `FFD9EAD3`. Centered both ways,
  wrap on, no borders.
- Bold on columns D, E, F only.
- Column H font color blue `FF0000FF` (it holds pasted URLs).
- Column widths, verbatim from the example: A 15.55, B 14.55, C 14.44,
  D 12.55, E 18.33, H 32.0, I 12.11, J 11.55, K 77.89, L 14.55, M 12.0,
  N 8.66. F and G are left at the default width, as in the example.
- No freeze panes — the example has none.

## Edge cases

- Missing source sheet → `ValueError` naming the sheet and listing the
  workbook's actual sheet names.
- Missing `Clip Name`/`Clip Duration` column → whatever `find_column_any`
  already raises, unchanged.
- Rows with a blank/whitespace-only clip name are skipped entirely and
  are not counted in `rows_in`.
- A blank or unparseable `Clip Duration` counts as 0 frames in the sum;
  it does not abort the run. The surviving row's own `CLIP DURATION`
  cell still shows the raw source value.
- An empty `3rd parties` sheet produces a `FU grid` with a header row
  and no data rows, and all-zero counts.

## Testing

`tests/test_fu_grid.py`, on a synthetic 3rd-parties fixture built in the
test (same spirit as `tests/conftest.py` — no client files):

- Exact-string grouping: two names differing only by extension stay two
  rows.
- `TOTAL USES` counts, and `rows_in` / `unique_clips` /
  `multi_use_clips` in the returned dict.
- `TOTAL DURATION` frame math across a group, including a frame carry
  (e.g. `00:00:00:20` + `00:00:00:10` at 25 fps = `00:00:00:05` in the
  next second).
- Single-use row's `TOTAL DURATION` equals its `CLIP DURATION`.
- Header text and order of all 13 columns, including the trailing space
  in `PREVIEW: `.
- Surviving row is the earliest `Sequence In`, and grid rows are ordered
  by it — verified with a fixture whose rows are deliberately out of
  order.
- Both output sheets exist, in order, and the `3rd parties` copy still
  has the source's row/column counts and values.
- Fill colors on a data row: `FFF4CCCC` in column A, `FFD9EAD3` in
  column I.
