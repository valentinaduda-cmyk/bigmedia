# Always-on standardized output styling

**Date:** 2026-09-10
**Status:** Approved, ready for implementation plan

## Problem

Column widths, data font, and header fill are currently opt-in per run
(`--autofit`, `--uniform-font`, `--uniform-header`, `--min-width`,
`--max-width`, and the matching web form checkboxes). In practice the team
wants them *always* applied and *always* the same, so every delivered
workbook reads consistently regardless of who ran the command or which
boxes they happened to tick. The options are friction with no upside.

## Goal

Remove the options. Make three styling behaviors unconditional and fixed
for the workbook-producing commands, applied to every **generated** sheet:

1. **Column widths** — each column sized to
   `max(longest data value, header text) + 2` characters, clamped to
   **[10, 60]**. Measured once per table and shared across sibling sheets
   so a column is the same width on every tab.
2. **Data font** — every data cell (row 2 downward) gets one font, sampled
   from the source's clip-name column at row 2 (Calibri 11 fallback if
   that cell carries no font). Keeps the file's house font, just uniform.
3. **Header row** — every non-empty header cell gets a **solid black fill
   (`FF000000`) and white bold text (`FFFFFF`)**, keeping the source
   header cell's font family and size. No longer "the dominant fill found
   in the source".

## Scope

Applies to **5 commands**: `sort`, `dedupe`, `group`, `compare`,
`fix-getty`.

**Excluded — left exactly as they are:**
- `fu-grid` — builds a hand-designed grid (deliberate widths, a
  `"PREVIEW: "` header with a trailing space, colours from a reference
  file).
- `getty-ids` — the Customer Declaration Form, styled to match Getty's own
  template (Lato, purple titles, fixed 20/10 column widths) so it can be
  sent to Getty unedited.

**Never styled — stays a verbatim copy of the input:**
- The `"Worksheet"` backup tab, and any other pass-through/backup sheet a
  command copies through unchanged (`copy_sheet_verbatim`).

## Decisions (from brainstorming)

- Every command in scope: the 5 above. fu-grid and getty-ids untouched.
- `"Worksheet"` / backup tabs stay byte-identical to the input — no width,
  font, or header change.
- Data font: sampled from the source body (clip-name column, row 2), not a
  fixed family.
- Header font: force white + bold, keep the source header's family/size;
  fill is fixed black.
- CLI flags and web form fields are **removed**, not kept as no-ops.

## Architecture

One shared entry point in `xlsx_utils.py` that all 5 commands call once,
after building their output workbook and before saving.

```
command builds wb_out (category / deduped / grouped / diff / getty sheets)
        │
        ▼
style_output_sheets(sheets, width_by=<"index"|"header">, data_font=<Font>)
        ├── measure widths  (index-keyed or header-name-keyed)
        ├── apply widths     (clamped [10, 60], + 2 padding)
        ├── apply uniform data font   (row 2+, every generated sheet)
        └── apply header style        (row 1: black fill, white bold)
        │
        ▼
wb_out.save(out_path)
```

**Why two width-keying modes:**
- `sort`, `dedupe`, `compare`'s row sheets, `fix-getty` — columns sit at
  fixed positions across sibling sheets → key widths by **column index**.
- `group` — grouping inserts columns mid-table, so the same header lands
  at different indices on different sheets → key widths by **header name**
  (existing `measure_header_widths` rationale).

`compare`'s `Summary` sheet keeps its bespoke fixed-width layout (cols A
= 24, others = 20) but its row-1 header still gets the black/white
treatment.

## Components

### `bigmedia/xlsx_utils.py`

**New:**
```python
HEADER_FILL = PatternFill("solid", fgColor="FF000000")
HEADER_FONT_COLOR = "FFFFFFFF"

def sample_data_font(ws, name_col_idx):
    """The row-2 font of the clip-name column, to standardize data cells
    on. Falls back to Calibri 11 when that cell has no font set."""

def apply_header_style(ws):
    """Row 1: every non-empty header cell gets HEADER_FILL and a white bold
    font, keeping the cell's existing family and size."""

def style_output_sheets(sheets, *, width_by, data_font):
    """The single styling pass every in-scope command runs on its finished
    output. `sheets` is the list of GENERATED worksheets (never the
    "Worksheet" backup — callers exclude it, and this also skips any sheet
    literally titled "Worksheet" as a guard). `width_by` is "index" or
    "header". Applies: shared column widths (clamped [10, 60]), the uniform
    `data_font` to row 2+, and apply_header_style to row 1."""
```

**Changed:**
- `measure_column_widths` / `measure_header_widths` — drop the
  `min_width` / `max_width` / `padding` parameters; hardcode to
  `AUTOFIT_MIN_WIDTH` (10), `AUTOFIT_MAX_WIDTH` (60), `AUTOFIT_PADDING`
  (2). Keep both functions (index vs header keying).
- `apply_uniform_data_font` — unchanged.
- `apply_column_widths` / `apply_header_widths` — unchanged.

**Deleted:**
- `dominant_header_fill` (and `DEFAULT_HEADER_FILL_COLOR` — folded into
  `HEADER_FILL`).
- `apply_uniform_header_fill` (replaced by `apply_header_style`, which also
  sets the font, not just the fill).

### `bigmedia/sort_workbook.py`

- `sort_workbook(...)` — drop `autofit`, `min_width`, `max_width`,
  `uniform_font` parameters (keep `category_order`, `categories`,
  `skip_categories`).
- After the category sheets are built, before `wb_out.save`, call
  `style_output_sheets(<category sheets>, width_by="index",
  data_font=sample_data_font(ws_src, name_col_idx))`.
- `count_categories` / `analyze_sort` — untouched (no output).

### `bigmedia/dedupe.py`

- `dedupe_workbook(src_path, out_path, name_column="Clip Name")` — no
  signature change (no styling params today). Now calls
  `style_output_sheets([<the Deduped and Duplicates worksheets>],
  width_by="index", data_font=sample_data_font(ws_src, name_col_idx))`
  before save. New behavior: deduped output is now styled where before it
  only copied the source header styling.

### `bigmedia/group_duplicates.py`

- `group_duplicates_workbook(...)` — drop `autofit`, `min_width`,
  `max_width`, `uniform_font`, `uniform_header` parameters.
- Replace the three conditional formatting blocks with one
  `style_output_sheets(formatted, width_by="header",
  data_font=_first_data_font(formatted) or <Calibri 11>)` call.
  `formatted` already excludes `"Worksheet"`.

### `bigmedia/compare_versions.py`

- `compare_workbooks(old_path, new_path, out_path, name_column="Clip Name",
  case_sensitive=False, unique=False)` — no signature change (it exposes no
  styling params). The module currently hardcodes `Font(bold=True)` header
  cells and copies/shifts source column widths.
- `_write_rows_sheet` puts a `Status` column at index 1, then the source
  columns shifted one right. Width keying is "header" (headers are unique,
  including `"Status"`).
- After building `Summary` + the `<cat> added` / `<cat> removed` sheets:
  - Run `apply_header_style` on every generated sheet's row 1 (including
    `Summary`).
  - Run `style_output_sheets(<the added/removed row sheets>,
    width_by="header", data_font=<sampled from the new file's body>)`.
  - `Summary` keeps its fixed column widths (A = 24, rest = 20); only its
    header row is restyled.

### `bigmedia/getty_split.py` (`fix_getty_split`)

- `fix_getty_split(old_path, new_path, out_path, name_column="Clip Name")`
  — no signature change (no styling params).
- After `_write_sheet` builds the Getty Videos / Getty Stills sheets,
  call `style_output_sheets(<those sheets>, width_by="index",
  data_font=<sampled from the source Getty sheet body>)`. Zebra row fills
  are preserved (data font sets the font, not the fill).

### `bigmedia/cli.py`

- Only `cmd_sort` and `cmd_group` pass these kwargs today. Remove
  `--autofit`, `--min-width`, `--max-width`, `--uniform-font`,
  `--uniform-header` from the shared options parser (the `for _p in (...)`
  loop around line 320), and drop the `autofit=`/`min_width=`/`max_width=`/
  `uniform_font=`/`uniform_header=` kwargs from the `sort_workbook(...)`
  and `group_duplicates_workbook(...)` calls in `cmd_sort` / `cmd_group`.
- Remove the `AUTOFIT_MIN_WIDTH` / `AUTOFIT_MAX_WIDTH` import from `cli.py`
  if nothing else there uses it after.

### `web/commands.py`

- Remove these `FieldSpec`s from `COMMANDS["sort"].fields`: `autofit`,
  `min_width`, `max_width`, `uniform_font`.
- Remove from `COMMANDS["group"].fields`: `autofit`, `min_width`,
  `max_width`, `uniform_font`, `uniform_header`.
- Remove the now-unused `AUTOFIT_MAX_WIDTH` / `AUTOFIT_MIN_WIDTH` import if
  unused after.

### `CLAUDE.md`

- Update the "Conventions" section: output styling (column widths, data
  font, black/white headers) is standardized and always applied to
  generated sheets; the `"Worksheet"` backup stays a verbatim copy.

## Testing

| File | Changes |
|---|---|
| `tests/test_xlsx_utils.py` | New: `style_output_sheets` (index vs header keying; [10,60] clamp; `+2` padding; black `FF000000` fill + white bold header keeping size; data font uniform on row 2+; a sheet titled `"Worksheet"` in the list is skipped). `sample_data_font` fallback to Calibri 11. |
| `tests/test_sort_workbook.py` | Rewrite the ~5 styling tests: remove `autofit=` / `min_width=` / `max_width=` / `uniform_font=` args (now unconditional). Assert every generated sheet's widths are within [10,60] and equal across sheets, headers are black-fill/white-bold, and the `"Worksheet"` tab is byte-identical to the input (unchanged assertion). Keep every classification / row-count / category test as-is. |
| `tests/test_group_duplicates.py` | Same: styling is always-on, header-keyed widths, black/white headers; drop the removed kwargs. Keep grouping-logic tests. |
| `tests/test_dedupe.py` | Add one styling assertion (Deduped/Duplicates headers black/white, widths in range). Keep dedup-key tests. |
| `tests/test_compare_versions.py` | Add one assertion: all generated sheets' header rows are black/white; added/removed sheets' widths in [10,60]; `Summary` keeps A=24. |
| `tests/test_getty_split.py` | Add one assertion: Getty Videos/Stills headers black/white, widths in range, zebra fills still present. |
| `tests/test_cli.py` | The removed flags no longer parse (`--autofit` etc. → argparse error). Existing command runs still succeed without them. |
| `tests/test_web_commands.py` | `sort` / `group` field-set assertions updated to the trimmed lists. |
| `tests/test_web_presets*.py` | Verify a saved preset that still carries a dropped key (e.g. `autofit`) is loaded without error — `parse_field` / `_build_kwargs` ignore keys with no matching `FieldSpec`. Adjust the 2 hits if they assert on those fields. |

Full `pytest` green.

## Out of scope

- Any change to `fu-grid` or `getty-ids` output.
- Restyling the `"Worksheet"` / backup tabs.
- Border, alignment, number-format, or row-height standardization —
  only widths, data font, and header fill/font colour.
- Migration of existing saved presets (dropped keys are simply ignored on
  load).
