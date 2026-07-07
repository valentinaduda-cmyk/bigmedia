# Getty Customer Declaration Form Output — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `build_getty_id_report()` emit a report styled like Getty's
real "Customer Declaration Form" — header fields, per-episode Video +
Stills blocks (reading both the `"Getty videos"` and `"Getty pics"`
sheets of each sorted input workbook), and dynamically-positioned
Instructions / Important Notes text boxes — instead of today's bare
single-block-per-file layout.

**Architecture:** All new styling (header fields, block stamping,
instructions/notes boxes) is implemented as Python constants and small
pure functions inside `bigmedia/getty_ids.py` — no binary template asset
is bundled. `build_getty_id_report()` orchestrates: write header once,
then loop input files writing a Video block immediately followed by a
Stills block (with column math tracked as it goes), then position the two
text boxes a fixed gap after the last block actually written.

**Tech Stack:** Python 3, openpyxl (already a dependency), pytest.

## Global Constraints

- `bigmedia/classify.py` is untouched by this work — not in scope.
- No binary `.xlsx` template is committed to the repo (confirmed decision
  — see `docs/superpowers/specs/2026-07-07-getty-declaration-form-design.md`).
- `project_name` is a **required** parameter/CLI arg (no default) —
  changes every job, a blank value would look broken on a real form.
- `broadcaster` defaults to `""` (blank) — matches the real example where
  this field was left empty.
- `production_company` defaults to `"KM Record a.s./Big Media"`,
  `rights` defaults to `"in perpetuity/worldwide/all media"`.
- Generated SUM/COUNTA formula ranges are always correct (`_FIRST_DATA_ROW`
  through the last actual data row) — the real example file has a
  copy-paste bug in its 2nd episode block (`N14` sums `N14:N3602` instead
  of `N15:N3602`); this bug is never replicated.
- `counts` returned by `build_getty_id_report()` is
  `{filename: {"video": n, "stills": n}}` — a breaking change from
  today's flat `{filename: n}`.
- Business logic stays I/O-light and testable without argparse, per
  `CLAUDE.md`: `build_getty_id_report(src, out, ...)` remains a plain
  function the CLI thinly wraps.
- Run `pytest` after every task and confirm the **whole** suite passes,
  not just the new/changed tests — this file has a history of one fix
  breaking previously-confirmed cases.

---

## Task 1: Header fields (Production Company / Project Name / Broadcaster / Rights)

Add the form's header block (rows 1–6) above the existing per-file
block, and make `project_name` a required parameter. Block content itself
is unchanged in this task — only shifted down 6 rows. No stills, no text
boxes yet (later tasks).

**Files:**
- Modify: `bigmedia/getty_ids.py:130-210` (style constants through
  `build_getty_id_report`)
- Test: `tests/test_getty_ids.py`

**Interfaces:**
- Produces: `build_getty_id_report(input_path, out_path, project_name, sheet_name="Getty videos", name_column="Clip Name", seconds_column="Seconds", min_seconds=5, production_company="KM Record a.s./Big Media", broadcaster="", rights="in perpetuity/worldwide/all media")` — `project_name` is now the 3rd positional arg, required.
- Produces: row constants `_FORM_TITLE_ROW=1`, `_COMPANY_ROW=2`, `_PROJECT_ROW=3`, `_BROADCASTER_ROW=4`, `_RIGHTS_ROW=5`, `_BLOCK_TITLE_ROW=7`, `_BLOCK_SUBTITLE_ROW=8`, `_BLOCK_HEADER_ROW=9`, `_BLOCK_FORMULA_ROW=10`, `_BLOCK_FIRST_DATA_ROW=11` — later tasks reference these names, not literal numbers.

- [ ] **Step 1: Update existing tests for the new row offsets and required `project_name`**

Replace the full contents of `tests/test_getty_ids.py` from
`def _make_sorted_workbook` onward (keep `CASES`, `test_extract_getty_id`,
`TITLE_CASES`, `test_clean_episode_title` unchanged above it) with:

```python
def _make_sorted_workbook(path, rows, stills_rows=None):
    """rows/stills_rows: list of (name, seconds) tuples. `seconds` mirrors
    the real convention: it's blank (None) on every row of a duplicate run
    except the last, which carries that clip's total-duration-in-seconds."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append(["Track", "Clip Name", "Seconds"])
    for name, seconds in rows:
        ws.append(["V1", name, seconds])
    if stills_rows is not None:
        ws_stills = wb.create_sheet(title="Getty pics")
        ws_stills.append(["Track", "Clip Name", "Seconds"])
        for name, seconds in stills_rows:
            ws_stills.append(["V1", name, seconds])
    ws2 = wb.create_sheet(title="AP")  # not a Getty sheet, must be ignored
    ws2.append(["Track", "Clip Name", "Seconds"])
    ws2.append(["V1", "BM1234_something.mxf", 10])
    wb.save(path)


def test_build_getty_id_report_writes_header_fields(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(
        str(in_dir), str(out_path), project_name="Top 10 Secrets of Technology",
        production_company="KM Record a.s./Big Media", broadcaster="Discovery",
        rights="in perpetuity/worldwide/all media",
    )

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=1, column=1).value == "Customer Declaration Form"
    assert ws.cell(row=2, column=1).value == "Production Company:"
    assert ws.cell(row=2, column=2).value == "KM Record a.s./Big Media"
    assert ws.cell(row=3, column=1).value == "Project Name:"
    assert ws.cell(row=3, column=2).value == "Top 10 Secrets of Technology"
    assert ws.cell(row=4, column=1).value == "Broadcaster:"
    assert ws.cell(row=4, column=2).value == "Discovery"
    assert ws.cell(row=5, column=1).value == "Rights Requested:"
    assert ws.cell(row=5, column=2).value == "in perpetuity/worldwide/all media"


def test_build_getty_id_report_header_field_defaults(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=2, column=2).value == "KM Record a.s./Big Media"
    assert ws.cell(row=4, column=2).value == ""
    assert ws.cell(row=5, column=2).value == "in perpetuity/worldwide/all media"


def test_build_getty_id_report_consolidates_multiple_files(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1_master_sorted.xlsx", [
        ("GettyImages-1001500162.mov", 6),
        ("GettyImages-1344-77.mov", 5),
    ])
    _make_sorted_workbook(in_dir / "EP2_master_sorted.xlsx", [
        ("GettyImages-mr_00108323.mov", 7),
    ])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    assert counts == {
        "EP1_master_sorted.xlsx": {"video": 2, "stills": 0},
        "EP2_master_sorted.xlsx": {"video": 1, "stills": 0},
    }

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]

    # First block: title, subtitle, Asset ID/Duration headers, formula row, then data.
    assert ws.cell(row=7, column=1).value == "EP1_master_sorted"
    assert ws.cell(row=8, column=1).value == "Getty Images Video"
    assert ws.cell(row=9, column=1).value == "Asset ID"
    assert ws.cell(row=9, column=2).value == "Duration"
    assert ws.cell(row=10, column=1).value == "=COUNTA(A11:A12)"
    assert ws.cell(row=10, column=2).value == "=SUM(B11:B12)"
    assert ws.cell(row=11, column=1).value == "1001500162"
    assert ws.cell(row=11, column=2).value == 6
    assert ws.cell(row=12, column=1).value == "1344-77"
    assert ws.cell(row=12, column=2).value == 5


def test_build_getty_id_report_dedupes_and_filters_by_seconds(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [
        ("GettyImages-1000000001.mov", 3),          # unique, too short -> excluded
        ("GettyImages-2000000002.mov", None),        # dup run: blank until last row
        ("GettyImages-2000000002.mov", None),
        ("GettyImages-2000000002.mov", 6),           # aggregate seconds -> one entry
        ("GettyImages-3000000003.mov", 5),           # unique, exactly at threshold -> included
    ])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 2, "stills": 0}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=11, column=1).value == "2000000002"
    assert ws.cell(row=11, column=2).value == 6
    assert ws.cell(row=12, column=1).value == "3000000003"
    assert ws.cell(row=12, column=2).value == 5


def test_build_getty_id_report_skips_garbled_seconds_cells(tmp_path):
    # A real file had a corrupted, non-numeric value in a "Seconds" cell —
    # must be treated as "doesn't qualify", not crash the whole batch.
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [
        ("GettyImages-1000000001.mov", "#REF!"),
        ("GettyImages-2000000002.mov", 6),
    ])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}


def test_build_getty_id_report_min_seconds_is_configurable(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1000000001.mov", 3)])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X", min_seconds=0)
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}


def test_build_getty_id_report_falls_back_to_name_column(tmp_path):
    # Real delivery files disagree: some use "Clip Name", others plain
    # "Name", even within the same batch. Neither should need -o/--name-column.
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1001500162.mov", 6)])

    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append([" Status", "Location", "Track", "Name", "Source", "Seconds"])
    ws.append([None, None, None, "GettyImages-1344-77.mov", "Getty", 5])
    wb.save(in_dir / "EP2.xlsx")

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {
        "EP1.xlsx": {"video": 1, "stills": 0},
        "EP2.xlsx": {"video": 1, "stills": 0},
    }

    wb_out = load_workbook(out_path)
    ws_out = wb_out["Getty IDs"]
    assert ws_out.cell(row=11, column=1).value == "1001500162"
    assert ws_out.cell(row=11, column=8).value == "1344-77"


def test_build_getty_id_report_ignores_blank_names(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    path = in_dir / "EP1.xlsx"
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append(["Track", "Clip Name", "Seconds"])
    ws.append(["V1", "GettyImages-1001500162.mov", 6])
    ws.append(["V1", None, None])
    wb.save(path)

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}
```

Note: Task 2 later changes the column layout to interleave Video/Stills
blocks per episode, which shifts EP2's block from column 4 to column 8.
`test_build_getty_id_report_falls_back_to_name_column`'s last assertion
(`column=4` in the code block above) gets updated to `column=8` as part
of Task 2, Step 1 — no action needed here in Task 1.

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_getty_ids.py -v`
Expected: FAIL — `TypeError: build_getty_id_report() missing 1 required positional argument: 'project_name'` (or similar) on every `build_getty_id_report(...)` call, since the current signature has no `project_name` param.

- [ ] **Step 3: Implement the header block and row-offset shift**

In `bigmedia/getty_ids.py`, replace everything from the `_TITLE_FONT = ...`
line through the end of the file with:

```python
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
    for file_path in files:
        rows = _read_getty_ids(file_path, sheet_name, name_column, seconds_column, min_seconds)
        counts[file_path.name] = {"video": len(rows), "stills": 0}

        title = clean_episode_title(file_path.name)
        _write_block(ws_out, col, title, "Getty Images Video", rows)
        col += 3  # two data columns + one blank spacer column

    ws_out.freeze_panes = f"A{_BLOCK_FIRST_DATA_ROW}"
    wb_out.save(out_path)
    return counts
```

This is a placeholder shape for `counts["stills"]` (`0` always, no stills
read yet) — Task 2 replaces the loop body to actually read and write the
stills block. Column stepping (`col += 3`) is unchanged from the original
so Task 1's tests (with `column=4` for EP2) pass as committed.

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_getty_ids.py -v`
Expected: PASS — all tests green.

- [ ] **Step 5: Run the full suite**

Run: `pytest -v`
Expected: PASS. (No other module imports `build_getty_id_report`'s old
signature except `cli.py`, which Task 4 updates — until then `cli.py`'s
call site is stale but nothing currently tests the CLI's `getty-ids`
subcommand, so this won't surface as a test failure. Confirm with
`pytest tests/test_cli.py -v` specifically that it's unaffected.)

- [ ] **Step 6: Commit**

```bash
git add bigmedia/getty_ids.py tests/test_getty_ids.py
git commit -m "feat(getty-ids): add Customer Declaration Form header fields"
```

---

## Task 2: Stills block (read "Getty pics", write Video+Stills pair per episode)

**Files:**
- Modify: `bigmedia/getty_ids.py` (add `stills_sheet_name` param and gap
  constants; replace the loop body in `build_getty_id_report`)
- Test: `tests/test_getty_ids.py`

**Interfaces:**
- Consumes: `_write_block(ws, col, title, subtitle, rows)` from Task 1 —
  called twice per file now (once per asset type).
- Produces: `build_getty_id_report(..., stills_sheet_name="Getty pics", ...)`
  — new keyword param inserted right after `sheet_name` in the signature.
- Produces: module-level `_INTRA_BLOCK_GAP = 1`, `_INTER_EPISODE_GAP = 2`
  constants and a `last_used_col` value tracked across the loop — Task 3
  reads this same tracked value to position the text boxes.

- [ ] **Step 1: Update tests for the new column layout and add stills-specific cases**

In `tests/test_getty_ids.py`:

1. Update `test_build_getty_id_report_falls_back_to_name_column`'s last
   assertion from `column=4` to `column=8` (per the note left at the end
   of Task 1):

```python
    assert ws_out.cell(row=11, column=1).value == "1001500162"
    assert ws_out.cell(row=11, column=8).value == "1344-77"
```

2. Add these new test functions at the end of the file:

```python
def test_build_getty_id_report_reads_stills_sheet(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1001500162.mov", 6)],
        stills_rows=[("GettyImages-2001500162.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 1}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]

    # Video block: columns 1-2 (A-B).
    assert ws.cell(row=7, column=1).value == "EP1"
    assert ws.cell(row=8, column=1).value == "Getty Images Video"
    assert ws.cell(row=11, column=1).value == "1001500162"
    assert ws.cell(row=11, column=2).value == 6

    # One blank spacer column (3 / C) between video and stills.
    assert ws.cell(row=7, column=3).value is None

    # Stills block: columns 4-5 (D-E).
    assert ws.cell(row=7, column=4).value == "EP1"
    assert ws.cell(row=8, column=4).value == "Getty Images Stills"
    assert ws.cell(row=9, column=4).value == "Asset ID"
    assert ws.cell(row=9, column=5).value == "Duration"
    assert ws.cell(row=10, column=4).value == "=COUNTA(D11:D11)"
    assert ws.cell(row=10, column=5).value == "=SUM(E11:E11)"
    assert ws.cell(row=11, column=4).value == "2001500162"
    assert ws.cell(row=11, column=5).value == 5


def test_build_getty_id_report_missing_stills_sheet_writes_empty_block(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", rows=[("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=8, column=4).value == "Getty Images Stills"
    assert ws.cell(row=10, column=4).value == "=COUNTA(D11:D11)"
    assert ws.cell(row=11, column=4).value is None


def test_build_getty_id_report_two_episodes_video_and_stills_column_math(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1000000001.mov", 6)],
        stills_rows=[("GettyImages-2000000001.jpg", 5)],
    )
    _make_sorted_workbook(
        in_dir / "EP2.xlsx",
        rows=[("GettyImages-1000000002.mov", 6)],
        stills_rows=[("GettyImages-2000000002.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    # EP1: video cols 1-2, stills cols 4-5. EP2: video cols 8-9, stills cols 11-12.
    assert ws.cell(row=11, column=1).value == "1000000001"
    assert ws.cell(row=11, column=4).value == "2000000001"
    assert ws.cell(row=11, column=8).value == "1000000002"
    assert ws.cell(row=11, column=11).value == "2000000002"

    # Two blank columns (6, 7 / F, G) between EP1's stills block and EP2's video block.
    assert ws.cell(row=7, column=6).value is None
    assert ws.cell(row=7, column=7).value is None
    assert ws.cell(row=7, column=8).value == "EP2"


def test_build_getty_id_report_custom_stills_sheet_name(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append(["Track", "Clip Name", "Seconds"])
    ws.append(["V1", "GettyImages-1001500162.mov", 6])
    ws_stills = wb.create_sheet(title="Getty Stills")  # non-default name
    ws_stills.append(["Track", "Clip Name", "Seconds"])
    ws_stills.append(["V1", "GettyImages-2001500162.jpg", 5])
    wb.save(in_dir / "EP1.xlsx")

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(
        str(in_dir), str(out_path), project_name="X", stills_sheet_name="Getty Stills",
    )
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 1}}
```

- [ ] **Step 2: Run tests to verify the new/changed ones fail**

Run: `pytest tests/test_getty_ids.py -v`
Expected: FAIL on the new stills tests (`build_getty_id_report() got an
unexpected keyword argument 'stills_sheet_name'`) and on the updated
`falls_back_to_name_column` assertion (column 8 is still blank, id is at
column 4).

- [ ] **Step 3: Implement stills reading and the two-block-per-episode layout**

In `bigmedia/getty_ids.py`, add these two constants right after
`_BLOCK_FIRST_DATA_ROW = 11`:

```python
_INTRA_BLOCK_GAP = 1    # blank columns between a Video block and its Stills block
_INTER_EPISODE_GAP = 2  # blank columns between one episode's Stills block and the next episode's Video block
```

Then replace `build_getty_id_report`'s signature and loop body with:

```python
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
```

(`last_used_col` is unused after the loop for now — Task 3 uses it to
position the text boxes. Leave it assigned.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_getty_ids.py -v`
Expected: PASS — all tests green.

- [ ] **Step 5: Run the full suite**

Run: `pytest -v`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add bigmedia/getty_ids.py tests/test_getty_ids.py
git commit -m "feat(getty-ids): read Getty pics sheet, write Video+Stills block pairs"
```

---

## Task 3: Instructions / Important Notes text boxes (dynamically positioned)

**Files:**
- Modify: `bigmedia/getty_ids.py`
- Test: `tests/test_getty_ids.py`

**Interfaces:**
- Consumes: `last_used_col` from Task 2's loop (the column of the last
  episode's Stills Duration cell).
- Produces: `_write_text_box(ws, col, width, header_text, body_text, header_border)` helper — no other task depends on this, it's called only from `build_getty_id_report`.

- [ ] **Step 1: Write failing tests**

Add to `tests/test_getty_ids.py`:

```python
_INSTRUCTIONS_TEXT_START = "Please divide content into appropriate asset type."
_NOTES_TEXT_START = "Upon receipt of the Proposed Usage Declaration form"


def test_build_getty_id_report_text_boxes_one_file(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", rows=[("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    # EP1 alone: video cols 1-2, stills cols 4-5, last_used_col=5.
    # box_start = 5 + 1 + 2 (BOX_GAP) = 8 (H).
    assert ws.cell(row=5, column=8).value == "Instructions:"
    assert ws.cell(row=7, column=8).value.startswith(_INSTRUCTIONS_TEXT_START)
    # notes_start = 8 + 6 (INSTR width) + 1 (inner gap) = 15 (O).
    assert ws.cell(row=5, column=15).value == "Important Notes on Licensing:"
    assert ws.cell(row=7, column=15).value.startswith(_NOTES_TEXT_START)


def test_build_getty_id_report_text_boxes_shift_with_two_files(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1000000001.mov", 6)],
        stills_rows=[("GettyImages-2000000001.jpg", 5)],
    )
    _make_sorted_workbook(
        in_dir / "EP2.xlsx",
        rows=[("GettyImages-1000000002.mov", 6)],
        stills_rows=[("GettyImages-2000000002.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    # EP1+EP2: video 8-9, stills 11-12, last_used_col=12.
    # box_start = 12 + 1 + 2 = 15 (O).
    assert ws.cell(row=5, column=15).value == "Instructions:"
    # notes_start = 15 + 6 + 1 = 22 (V).
    assert ws.cell(row=5, column=22).value == "Important Notes on Licensing:"


def test_build_getty_id_report_text_box_merge_ranges(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", rows=[("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    merged = {str(r) for r in ws.merged_cells.ranges}
    assert "H5:M6" in merged    # Instructions header, 6 cols wide, 2 rows tall
    assert "H7:M32" in merged   # Instructions body, 6 cols wide, 26 rows tall
    assert "O5:V6" in merged    # Notes header, 8 cols wide, 2 rows tall
    assert "O7:V32" in merged   # Notes body, 8 cols wide, 26 rows tall
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_getty_ids.py -v`
Expected: FAIL — the new test functions raise `AssertionError` (cells are
`None`) since no text box code exists yet.

- [ ] **Step 3: Implement the text boxes**

In `bigmedia/getty_ids.py`, add these constants after `_INTER_EPISODE_GAP`:

```python
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
```

Then, in `build_getty_id_report`, replace:

```python
    ws_out.freeze_panes = f"A{_BLOCK_FIRST_DATA_ROW}"
    wb_out.save(out_path)
    return counts
```

with:

```python
    box_start_col = last_used_col + 1 + _BOX_GAP
    _write_text_box(ws_out, box_start_col, _INSTRUCTIONS_WIDTH, "Instructions:", _INSTRUCTIONS_TEXT, _DASHED_BOTTOM)
    notes_start_col = box_start_col + _INSTRUCTIONS_WIDTH + _BOX_INNER_GAP
    _write_text_box(ws_out, notes_start_col, _NOTES_WIDTH, "Important Notes on Licensing:", _NOTES_TEXT, _DOTTED_BOTTOM)

    ws_out.freeze_panes = f"A{_BLOCK_FIRST_DATA_ROW}"
    wb_out.save(out_path)
    return counts
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_getty_ids.py -v`
Expected: PASS — all tests green.

- [ ] **Step 5: Run the full suite**

Run: `pytest -v`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add bigmedia/getty_ids.py tests/test_getty_ids.py
git commit -m "feat(getty-ids): add dynamically-positioned Instructions/Notes boxes"
```

---

## Task 4: CLI — new flags, nested counts summary

**Files:**
- Modify: `bigmedia/cli.py:98-149`

**Interfaces:**
- Consumes: `build_getty_id_report(input_path, out_path, project_name, sheet_name, stills_sheet_name, name_column, seconds_column, min_seconds, production_company, broadcaster, rights)` from Task 3 — all params now exist.

- [ ] **Step 1: Update `cmd_getty_ids` and the `getty-ids` argparse subparser**

In `bigmedia/cli.py`, replace the `cmd_getty_ids` function (currently
lines 98-115) with:

```python
def cmd_getty_ids(args):
    input_path = Path(args.input)
    if args.output:
        out_path = args.output
    elif input_path.is_dir():
        out_path = str(input_path / "getty_ids.xlsx")
    else:
        out_path = _default_output(input_path, "getty_ids")

    counts = build_getty_id_report(
        str(input_path), out_path, args.project_name,
        sheet_name=args.sheet, stills_sheet_name=args.stills_sheet,
        name_column=args.name_column, seconds_column=args.seconds_column,
        min_seconds=args.min_seconds,
        production_company=args.production_company,
        broadcaster=args.broadcaster, rights=args.rights,
    )
    print(f"Wrote {out_path}")
    for name, n in counts.items():
        print(f"  {name}: {n['video']} video, {n['stills']} stills")
```

Then replace the `p_getty` block (currently lines 142-149) with:

```python
    p_getty = subparsers.add_parser("getty-ids", help="Extract Getty clip ids from several sorted workbooks into a Customer Declaration Form report")
    p_getty.add_argument("input", help="Path to a folder of sorted .xlsx files (or a single file)")
    p_getty.add_argument("-o", "--output", help="Output report path (default: <input folder>/getty_ids.xlsx)")
    p_getty.add_argument("--project-name", required=True, help="Project Name for the form's header (required)")
    p_getty.add_argument("--production-company", default="KM Record a.s./Big Media", help='Production Company for the form\'s header (default: "KM Record a.s./Big Media")')
    p_getty.add_argument("--broadcaster", default="", help="Broadcaster for the form's header (default: blank)")
    p_getty.add_argument("--rights", default="in perpetuity/worldwide/all media", help='Rights Requested for the form\'s header (default: "in perpetuity/worldwide/all media")')
    p_getty.add_argument("--sheet", default="Getty videos", help='Sheet to read video clip names from (default: "Getty videos")')
    p_getty.add_argument("--stills-sheet", default="Getty pics", help='Sheet to read stills clip names from (default: "Getty pics")')
    p_getty.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_getty.add_argument("--seconds-column", default="Seconds", help='Header of the total-duration-in-seconds column (default: "Seconds")')
    p_getty.add_argument("--min-seconds", type=float, default=5, help="Only include clips whose total duration is at least this many seconds (default: 5)")
    p_getty.set_defaults(func=cmd_getty_ids)
```

Also update the module docstring at the top of `bigmedia/cli.py` (around
line 19, `"getty-ids" reads the "Getty videos" sheet...`) to:

```
"getty-ids" reads the "Getty videos" and "Getty pics" sheets of every
input workbook (customizable via --sheet/--stills-sheet), keeps only
unique clips with total duration >= --min-seconds, extracts each clip's
Getty id from its filename, and writes one Customer Declaration Form-
styled report (--project-name is required; --production-company,
--broadcaster, --rights fill the rest of the form's header).
```

- [ ] **Step 2: Run the full suite**

Run: `pytest -v`
Expected: PASS — `test_cli.py` doesn't exercise the `getty-ids`
subcommand today, so this is a smoke check that nothing else broke.

- [ ] **Step 3: Manually verify `--help` output**

Run: `py -c "from bigmedia.cli import main; main(['getty-ids', '--help'])"`
Expected: prints usage showing `--project-name` (marked required),
`--production-company`, `--broadcaster`, `--rights`, `--sheet`,
`--stills-sheet`, `--name-column`, `--seconds-column`, `--min-seconds`.

- [ ] **Step 4: Commit**

```bash
git add bigmedia/cli.py
git commit -m "feat(cli): add getty-ids header-field flags and stills sheet option"
```

---

## Task 5: Verify against a real file

**Files:** none modified — this is a manual verification task.

- [ ] **Step 1: Run against the real EP1 file**

Run:
```bash
py -c "
from bigmedia.cli import main
main(['getty-ids',
      r'C:\Users\valen\OneDrive\Documents\DEV\bigmedia\data\in\report_in_5.07\EP1 - Eiffel Tower (kopie)_grouped.xlsx',
      '-o', r'C:\Users\valen\OneDrive\Documents\DEV\bigmedia\data\out\getty_ids_ep1_test.xlsx',
      '--project-name', 'Top 10 Secrets of Technology'])
"
```
Expected: prints `Wrote ...` and a summary line like
`EP1 - Eiffel Tower (kopie)_grouped.xlsx: 52 video, N stills`.

- [ ] **Step 2: Open the output and sanity-check against the attached example**

Report back to Valentina, comparing counts against the old (Task-0)
baseline run (52 video ids from `"Getty videos"` alone) and the real
`Customer Declaration Form.xlsx` layout:
- header fields show the passed-in Project Name and default Production
  Company/Rights
- Video block still shows 52 ids
- Stills block shows however many qualify from `"Getty pics"`
- Instructions/Important Notes boxes appear immediately right of the
  single episode's blocks (not at fixed columns T/Z)

- [ ] **Step 3: Delete the manual-verification output file (not a committed artifact)**

```bash
rm "C:\Users\valen\OneDrive\Documents\DEV\bigmedia\data\out\getty_ids_ep1_test.xlsx"
```

(`data/out/` is gitignored per `CLAUDE.md`, but no need to leave scratch
files lying around.)

- [ ] **Step 4: No commit** — this task is verification only, nothing to commit.
