# Always-On Output Styling Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make column widths (clamped 10–60), a uniform data font, and black-fill/white-bold headers unconditional for the `sort`, `dedupe`, `group`, `compare`, and `fix-getty` commands, and delete the options that used to control them.

**Architecture:** One shared entry point `style_output_sheets(sheets, *, width_by, data_font)` in `bigmedia/xlsx_utils.py` that every in-scope command calls once on its generated sheets before saving. The `"Worksheet"` backup tab and the `fu-grid` / `getty-ids` commands are untouched.

**Tech Stack:** Python 3.9+, openpyxl, pytest. FastAPI + Jinja for the web form field removal.

**Spec:** `docs/superpowers/specs/2026-09-10-always-on-output-styling-design.md`

## Global Constraints

- Styling applies only to **generated** sheets. The `"Worksheet"` backup and any `copy_sheet_verbatim` pass-through sheet stays byte-identical to the input. `style_output_sheets` also skips any sheet titled exactly `"Worksheet"` as a guard.
- Header fill is fixed `PatternFill("solid", fgColor="FF000000")`; header font is forced white (`FFFFFFFF`) + `bold=True`, keeping the source header cell's family and size.
- Column widths: `min(max(longest_value_or_header + 2, 10), 60)`. Constants `AUTOFIT_MIN_WIDTH=10`, `AUTOFIT_MAX_WIDTH=60`, `AUTOFIT_PADDING=2` stay in `xlsx_utils.py`.
- Data font: sampled from the source (clip-name column row 2, or first data cell), Calibri 11 fallback.
- `fu-grid` and `getty-ids` output: **no change**.
- No new runtime dependencies. Commit style `feat: …` / `feat(web): …` / `refactor: …` / `test: …`; end every commit body with `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>`.
- Full `pytest` stays green (435 passing before this plan).
- Run tests with `python -m pytest`.

---

### Task 1: Styling foundation in `xlsx_utils.py`

**Files:**
- Modify: `bigmedia/xlsx_utils.py`
- Test: `tests/test_xlsx_utils.py`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `HEADER_FILL: PatternFill` — solid `FF000000`.
  - `apply_header_style(ws)` — row 1: every non-empty header cell gets `HEADER_FILL` and a white (`FFFFFFFF`) bold font that keeps the cell's existing `name`/`size`.
  - `sample_data_font(ws, col_idx) -> Font` — `copy.copy(ws.cell(row=2, column=col_idx).font)`, or `Font(name="Calibri", size=11)` when there's no row 2 or the cell has no explicit font.
  - `first_data_font(worksheets) -> Font` — first non-`None` data cell (row 2+) font scanning the given sheets, same Calibri-11 fallback. (Moved and generalized from `group_duplicates._first_data_font`, which had no fallback.)
  - `style_output_sheets(sheets, *, width_by, data_font) -> None` — for each sheet in `sheets` whose title is not `"Worksheet"`: measure shared widths (`width_by="index"` → `measure_column_widths` keyed by column position; `width_by="header"` → `measure_header_widths` keyed by header name), apply them (`apply_column_widths` / `apply_header_widths`), apply `data_font` to every data cell via `apply_uniform_data_font`, and run `apply_header_style`. Widths are measured once across all the sheets and shared.
  - `measure_column_widths(rows, headers=None)` and `measure_header_widths(worksheets)` — the `min_width` / `max_width` / `padding` parameters are **removed**; the functions use the module constants directly.
- Deletes: `dominant_header_fill`, `DEFAULT_HEADER_FILL_COLOR`, `apply_uniform_header_fill`.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_xlsx_utils.py` additions (the file exists from an earlier feature — append):

```python
import copy
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from bigmedia.xlsx_utils import (
    HEADER_FILL, apply_header_style, sample_data_font, first_data_font,
    style_output_sheets, AUTOFIT_MIN_WIDTH, AUTOFIT_MAX_WIDTH,
)


def _sheet(wb, title, rows):
    ws = wb.create_sheet(title=title)
    for row in rows:
        ws.append(row)
    return ws


def test_header_fill_is_solid_black():
    assert HEADER_FILL.patternType == "solid"
    assert HEADER_FILL.fgColor.rgb == "FF000000"


def test_apply_header_style_black_fill_white_bold_keeps_size():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Clip Name"
    ws["A1"].font = Font(name="Arial", size=14)
    ws["B1"] = None
    apply_header_style(ws)
    assert ws["A1"].fill.fgColor.rgb == "FF000000"
    assert ws["A1"].font.bold is True
    assert ws["A1"].font.color.rgb == "FFFFFFFF"
    assert ws["A1"].font.name == "Arial" and ws["A1"].font.size == 14


def test_sample_data_font_reads_row2_or_falls_back():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "h"
    ws["A2"] = "v"
    ws["A2"].font = Font(name="Verdana", size=9)
    assert sample_data_font(ws, 1).name == "Verdana"
    empty = Workbook().active
    empty["A1"] = "h"
    assert sample_data_font(empty, 1).name == "Calibri"


def test_first_data_font_scans_sheets_with_fallback():
    wb = Workbook()
    a = _sheet(wb, "A", [["h"], []])          # header only, no data
    b = _sheet(wb, "B", [["h"], ["x"]])
    b["A2"].font = Font(name="Tahoma", size=8)
    assert first_data_font([a, b]).name == "Tahoma"
    assert first_data_font([a]).name == "Calibri"


def test_style_output_sheets_index_keyed_clamped_and_uniform():
    wb = Workbook()
    wb.remove(wb.active)
    s1 = _sheet(wb, "AP", [["Clip Name", "Notes"], ["x" * 200, "n"]])
    s2 = _sheet(wb, "Getty Videos", [["Clip Name", "Notes"], ["y", "n"]])
    style_output_sheets([s1, s2], width_by="index", data_font=Font(name="Calibri", size=11))
    w1 = s1.column_dimensions["A"].width
    assert AUTOFIT_MIN_WIDTH <= w1 <= AUTOFIT_MAX_WIDTH
    assert s1.column_dimensions["A"].width == s2.column_dimensions["A"].width  # shared
    assert s1["A1"].fill.fgColor.rgb == "FF000000"
    assert s1["A2"].font.name == "Calibri"


def test_style_output_sheets_header_keyed_matches_by_name():
    wb = Workbook()
    wb.remove(wb.active)
    # same header at different column indices
    s1 = _sheet(wb, "S1", [["Clip Name", "Dur"], ["aaaa", 1]])
    s2 = _sheet(wb, "S2", [["Dur", "Clip Name"], [1, "aaaa"]])
    style_output_sheets([s1, s2], width_by="header", data_font=Font(name="Calibri", size=11))
    assert s1.column_dimensions["A"].width == s2.column_dimensions["B"].width  # "Clip Name" both


def test_style_output_sheets_skips_worksheet_tab():
    wb = Workbook()
    wb.remove(wb.active)
    ws = _sheet(wb, "Worksheet", [["Clip Name"], ["x"]])
    ws["A1"].fill = PatternFill()  # no fill
    style_output_sheets([ws], width_by="index", data_font=Font(name="Calibri", size=11))
    assert ws["A1"].fill.patternType is None  # untouched
```

- [ ] **Step 2: Run — verify they fail**

Run: `python -m pytest tests/test_xlsx_utils.py -v`
Expected: FAIL — `ImportError` for the new names.

- [ ] **Step 3: Implement**

In `bigmedia/xlsx_utils.py`:

1. Add `from openpyxl.styles import Font, PatternFill` to the imports.

2. Replace the `DEFAULT_HEADER_FILL_COLOR` / `dominant_header_fill` / `apply_uniform_header_fill` block with:

```python
HEADER_FILL = PatternFill("solid", fgColor="FF000000")
_HEADER_FONT_COLOR = "FFFFFFFF"
_FALLBACK_DATA_FONT = Font(name="Calibri", size=11)


def apply_header_style(ws):
    """Row 1: every non-empty header cell gets the fixed black fill and a
    white bold font, keeping the cell's own family and size."""
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        if cell.value is None:
            continue
        cell.fill = copy.copy(HEADER_FILL)
        f = cell.font
        cell.font = Font(name=f.name, size=f.size, bold=True, color=_HEADER_FONT_COLOR)


def sample_data_font(ws, col_idx):
    """The row-2 font of column `col_idx`, to standardize data cells on.
    Calibri 11 when there is no row 2 or the cell carries no font."""
    if ws.max_row < 2:
        return copy.copy(_FALLBACK_DATA_FONT)
    f = ws.cell(row=2, column=col_idx).font
    if f is None or f.name is None:
        return copy.copy(_FALLBACK_DATA_FONT)
    return copy.copy(f)


def first_data_font(worksheets):
    """First non-empty data cell (row 2+) font across the given sheets;
    Calibri 11 fallback."""
    for ws in worksheets:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None and cell.font is not None and cell.font.name is not None:
                    return copy.copy(cell.font)
    return copy.copy(_FALLBACK_DATA_FONT)


def style_output_sheets(sheets, *, width_by, data_font):
    """The single styling pass every in-scope command runs on its finished
    output sheets. Skips any sheet titled "Worksheet". `width_by` is
    "index" (columns fixed across sheets) or "header" (columns shift, key
    widths by header name). Applies shared column widths clamped to
    [10, 60], the uniform `data_font` to row 2+, and apply_header_style to
    row 1."""
    targets = [ws for ws in sheets if ws.title != "Worksheet"]
    if not targets:
        return
    if width_by == "index":
        widths = measure_column_widths(
            (
                [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                for ws in targets
                for r in range(2, ws.max_row + 1)
            ),
            headers=[
                ws.cell(row=1, column=c).value
                for ws in targets
                for c in range(1, ws.max_column + 1)
            ],
        )
        for ws in targets:
            apply_column_widths(ws, widths)
    elif width_by == "header":
        widths = measure_header_widths(targets)
        for ws in targets:
            apply_header_widths(ws, widths)
    else:
        raise ValueError(f"width_by must be 'index' or 'header', got {width_by!r}")
    for ws in targets:
        apply_uniform_data_font(ws, data_font)
        apply_header_style(ws)
```

3. `measure_column_widths` — change signature to `def measure_column_widths(rows, headers=None):` and inside use `AUTOFIT_MIN_WIDTH`, `AUTOFIT_MAX_WIDTH`, `AUTOFIT_PADDING` directly.

4. `measure_header_widths` — change signature to `def measure_header_widths(worksheets):` and use the constants directly.

- [ ] **Step 4: Run — verify pass**

Run: `python -m pytest tests/test_xlsx_utils.py -v`
Expected: PASS.

- [ ] **Step 5: Run full suite (expect breakage in command tests — that's fine, later tasks fix them)**

Run: `python -m pytest -q 2>&1 | tail -20`
Expected: `tests/test_sort_workbook.py` and `tests/test_group_duplicates.py` fail (they call the removed params / functions). `tests/test_xlsx_utils.py` and everything unrelated passes. **Record which tests fail** — Tasks 2 and 4 must clear exactly those.

- [ ] **Step 6: Commit**

```bash
git add bigmedia/xlsx_utils.py tests/test_xlsx_utils.py
git commit -m "$(cat <<'EOF'
feat: add style_output_sheets and fixed black/white header styling

One entry point for the always-on output styling: shared column widths
(clamped 10-60), a uniform data font, and a fixed black-fill/white-bold
header row. Drops the parameterized min/max width and the
dominant-fill header helpers.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 2: Wire `sort_workbook`

**Files:**
- Modify: `bigmedia/sort_workbook.py`
- Test: `tests/test_sort_workbook.py`

**Interfaces:**
- Consumes: `style_output_sheets`, `sample_data_font` (Task 1).
- Produces: `sort_workbook(src_path, out_path, name_column="Clip Name", category_order=None, categories=None, skip_categories=None)` — the `autofit`, `min_width`, `max_width`, `uniform_font` parameters are gone. Every generated category sheet is styled unconditionally.

- [ ] **Step 1: Rewrite the styling tests**

In `tests/test_sort_workbook.py`, replace the styling tests (`test_sort_workbook_autofit_widths_are_uniform_across_sheets`, `test_sort_workbook_autofit_respects_min_and_max`, `test_sort_workbook_autofit_fits_the_longest_value_in_the_column`, `test_sort_workbook_uniform_font_on_data_cells`, and keep `test_sort_workbook_worksheet_backup_is_untouched_by_formatting` but drop its kwargs) with:

```python
def test_sort_output_sheets_are_always_styled(sample_master, tmp_path):
    from openpyxl import load_workbook
    out = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out))
    wb = load_workbook(str(out))
    cat_sheets = [t for t in wb.sheetnames if t != "Worksheet"]
    assert cat_sheets
    ref = None
    for t in cat_sheets:
        ws = wb[t]
        # header row: black fill, white bold
        h = ws.cell(row=1, column=1)
        assert h.fill.fgColor.rgb == "FF000000"
        assert h.font.bold and h.font.color.rgb == "FFFFFFFF"
        # widths in range and shared across sheets
        w = ws.column_dimensions["A"].width
        assert 10 <= w <= 60
        ref = ref or w
        assert ws.column_dimensions["A"].width == ref
        # uniform data font (row 2+ all one name), when the sheet has data
        fonts = {ws.cell(row=r, column=1).font.name
                 for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=1).value}
        assert len(fonts) <= 1


def test_sort_worksheet_backup_stays_verbatim(sample_master, tmp_path):
    from openpyxl import load_workbook
    out = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out))
    wb = load_workbook(str(out))
    src = load_workbook(str(sample_master)).active
    bak = wb["Worksheet"]
    # header fill NOT forced to black on the backup
    assert bak.cell(row=1, column=1).fill.fgColor.rgb != "FF000000" or \
           src.cell(row=1, column=1).fill.fgColor.rgb == "FF000000"
    # values identical
    for r in range(1, src.max_row + 1):
        for c in range(1, src.max_column + 1):
            assert bak.cell(row=r, column=c).value == src.cell(row=r, column=c).value
```

- [ ] **Step 2: Run — verify the OLD tests fail / new ones error on kwargs**

Run: `python -m pytest tests/test_sort_workbook.py -v`
Expected: FAIL — old tests pass removed kwargs; new tests fail because styling isn't applied yet.

- [ ] **Step 3: Implement**

In `bigmedia/sort_workbook.py`:

1. Imports: drop `measure_column_widths`, `apply_column_widths`, `apply_uniform_data_font`, `AUTOFIT_MIN_WIDTH`, `AUTOFIT_MAX_WIDTH` if now unused; add `style_output_sheets`, `sample_data_font`.

2. Signature:
```python
def sort_workbook(src_path, out_path, name_column="Clip Name", category_order=None,
                  categories=None, skip_categories=None):
```

3. Delete the `autofit_widths` block (the `autofit_widths = None` / `if autofit:` lines).

4. In the per-category loop, replace the width + font block:
```python
        if autofit_widths:
            apply_column_widths(ws_out, autofit_widths)
        else:
            for col_letter, width in col_widths.items():
                ws_out.column_dimensions[col_letter].width = width
        if uniform_font:
            apply_uniform_data_font(ws_out, col_font[name_col_idx - 1])
        ws_out.freeze_panes = "A2"
```
with just:
```python
        for col_letter, width in col_widths.items():
            ws_out.column_dimensions[col_letter].width = width
        ws_out.freeze_panes = "A2"
```
(the `col_widths` seed is harmless; `style_output_sheets` overwrites it.)

5. Before `wb_out.save(out_path)`:
```python
    style_output_sheets(
        [wb_out[c[:31]] for c in category_order],
        width_by="index",
        data_font=sample_data_font(ws_src, name_col_idx),
    )
```

- [ ] **Step 4: Run — sort tests + full suite**

Run: `python -m pytest tests/test_sort_workbook.py -q` → PASS.
Run: `python -m pytest -q 2>&1 | tail -15` → only `test_group_duplicates.py` (and possibly `test_cli.py` / `test_web_commands.py`) still failing.

- [ ] **Step 5: Commit**

```bash
git add bigmedia/sort_workbook.py tests/test_sort_workbook.py
git commit -m "$(cat <<'EOF'
feat: always style sort output sheets, drop the formatting options

sort_workbook no longer takes autofit / min_width / max_width /
uniform_font -- every category sheet is width-fitted, uniform-fonted and
black/white-headed unconditionally. Worksheet backup stays verbatim.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 3: Wire `dedupe_workbook`

**Files:**
- Modify: `bigmedia/dedupe.py`
- Test: `tests/test_dedupe.py`

**Interfaces:**
- Consumes: `style_output_sheets`, `sample_data_font` (Task 1).
- Produces: `dedupe_workbook(src_path, out_path, name_column="Clip Name")` — unchanged signature; the `Deduped` and `Duplicates` sheets are now styled.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_dedupe.py`:

```python
def test_dedupe_output_sheets_are_styled(sample_master, tmp_path):
    from openpyxl import load_workbook
    out = tmp_path / "out.xlsx"
    dedupe_workbook(str(sample_master), str(out))
    wb = load_workbook(str(out))
    for t in ("Deduped", "Duplicates"):
        ws = wb[t]
        h = ws.cell(row=1, column=1)
        assert h.fill.fgColor.rgb == "FF000000"
        assert h.font.bold and h.font.color.rgb == "FFFFFFFF"
        assert 10 <= ws.column_dimensions["A"].width <= 60
    # backup untouched
    bak = wb["Worksheet"]
    src = load_workbook(str(sample_master)).active
    assert bak.cell(row=1, column=1).value == src.cell(row=1, column=1).value
```

- [ ] **Step 2: Run — verify it fails**

Run: `python -m pytest tests/test_dedupe.py -k styled -v`
Expected: FAIL — headers not black.

- [ ] **Step 3: Implement**

In `bigmedia/dedupe.py`:
1. Imports: add `style_output_sheets`, `sample_data_font` to the `from .xlsx_utils import ...` line.
2. After the `for title, rows in [("Deduped", ...), ("Duplicates", ...)]` loop, before `wb_out.save(out_path)`:
```python
    style_output_sheets(
        [wb_out["Deduped"], wb_out["Duplicates"]],
        width_by="index",
        data_font=sample_data_font(ws_src, name_col_idx),
    )
```
3. The inner `for col_letter, width in col_widths.items(): ...` seed can stay (harmless, overwritten).

- [ ] **Step 4: Run**

Run: `python -m pytest tests/test_dedupe.py -q` → PASS.

- [ ] **Step 5: Commit**

```bash
git add bigmedia/dedupe.py tests/test_dedupe.py
git commit -m "$(cat <<'EOF'
feat: style dedupe output sheets

Deduped and Duplicates now get the standard width-fit, uniform font and
black/white header. Worksheet backup unchanged.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 4: Wire `group_duplicates_workbook`

**Files:**
- Modify: `bigmedia/group_duplicates.py`
- Test: `tests/test_group_duplicates.py`

**Interfaces:**
- Consumes: `style_output_sheets`, `first_data_font` (Task 1).
- Produces: `group_duplicates_workbook(src_path, out_path, sheets=None, name_column="Clip Name", duration_column="Clip Duration", fps=25)` — the `autofit`, `min_width`, `max_width`, `uniform_font`, `uniform_header` parameters are gone. Grouped sheets styled unconditionally (widths keyed by header name).

- [ ] **Step 1: Rewrite the styling tests**

In `tests/test_group_duplicates.py`, drop the removed kwargs from every call and replace the autofit/uniform_header/uniform_font assertions with one:

```python
def test_group_output_is_always_styled(sample_sorted_master, tmp_path):
    from openpyxl import load_workbook
    out = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out))
    wb = load_workbook(str(out))
    styled = [t for t in wb.sheetnames if t != "Worksheet"]
    assert styled
    for t in styled:
        ws = wb[t]
        # find the Clip Name column, assert its header is black/white and width in range
        name_c = next(c for c in range(1, ws.max_column + 1)
                      if ws.cell(row=1, column=c).value == "Clip Name")
        h = ws.cell(row=1, column=name_c)
        assert h.fill.fgColor.rgb == "FF000000"
        assert h.font.bold and h.font.color.rgb == "FFFFFFFF"
        from openpyxl.utils import get_column_letter
        assert 10 <= ws.column_dimensions[get_column_letter(name_c)].width <= 60
```

(Keep all grouping-logic tests. Check the fixture name — `sample_sorted_master` from `conftest.py`.)

- [ ] **Step 2: Run — verify failures**

Run: `python -m pytest tests/test_group_duplicates.py -v`
Expected: FAIL — removed kwargs and missing styling.

- [ ] **Step 3: Implement**

In `bigmedia/group_duplicates.py`:
1. Imports: from `.xlsx_utils` drop `measure_header_widths`, `apply_header_widths`, `dominant_header_fill`, `apply_uniform_header_fill`, `apply_uniform_data_font`, `AUTOFIT_MIN_WIDTH`, `AUTOFIT_MAX_WIDTH`; add `style_output_sheets`, `first_data_font`.
2. Signature: drop `autofit`, `min_width`, `max_width`, `uniform_font`, `uniform_header`.
3. Replace the three `if autofit ... / if uniform_header ... / if uniform_font ...` blocks with:
```python
    formatted = [wb_out[t] for t in wb_out.sheetnames if t != "Worksheet"]
    if formatted:
        style_output_sheets(formatted, width_by="header", data_font=first_data_font(formatted))
```
4. Delete the now-unused local `_first_data_font` function (its callers are gone; `first_data_font` from xlsx_utils replaces it).

- [ ] **Step 4: Run**

Run: `python -m pytest tests/test_group_duplicates.py -q` → PASS.
Run: `python -m pytest -q 2>&1 | tail -15` → only `test_cli.py` / `test_web_commands.py` / `test_web_presets*.py` possibly still failing (Task 7).

- [ ] **Step 5: Commit**

```bash
git add bigmedia/group_duplicates.py tests/test_group_duplicates.py
git commit -m "$(cat <<'EOF'
feat: always style group output, drop the formatting options

group_duplicates_workbook no longer takes autofit / min_width /
max_width / uniform_font / uniform_header. Grouped sheets get
header-keyed width-fit, uniform font and black/white headers every run.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 5: Wire `compare_workbooks`

**Files:**
- Modify: `bigmedia/compare_versions.py`
- Test: `tests/test_compare_versions.py`

**Interfaces:**
- Consumes: `style_output_sheets`, `apply_header_style`, `sample_data_font` (Task 1).
- Produces: `compare_workbooks(old_path, new_path, out_path, name_column="Clip Name", case_sensitive=False, unique=False)` — unchanged signature. `Summary` header row → black/white (keeps its fixed A=24 / rest=20 widths). The `<cat> added` / `<cat> removed` row sheets → full styling, widths keyed by header name.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_compare_versions.py` (use its existing fixture that produces old + new masters; check the file for the helper):

```python
def test_compare_output_headers_are_black_white(compare_pair, tmp_path):
    from openpyxl import load_workbook
    old_p, new_p = compare_pair
    out = tmp_path / "cmp.xlsx"
    compare_workbooks(str(old_p), str(new_p), str(out))
    wb = load_workbook(str(out))
    for t in wb.sheetnames:
        ws = wb[t]
        h = ws.cell(row=1, column=1)
        assert h.fill.fgColor.rgb == "FF000000", t
        assert h.font.bold and h.font.color.rgb == "FFFFFFFF", t
    # a row sheet has widths in range
    row_sheets = [t for t in wb.sheetnames if t.endswith("added") or t.endswith("removed")]
    if row_sheets:
        ws = wb[row_sheets[0]]
        assert 10 <= ws.column_dimensions["A"].width <= 60
    # Summary keeps its fixed first-column width
    assert wb["Summary"].column_dimensions["A"].width == 24
```

(If there is no `compare_pair` fixture, build the pair inline with `openpyxl` the way the existing compare tests do.)

- [ ] **Step 2: Run — verify it fails**

Run: `python -m pytest tests/test_compare_versions.py -k black_white -v`
Expected: FAIL — `Font(bold=True)` headers have no fill.

- [ ] **Step 3: Implement**

In `bigmedia/compare_versions.py`:
1. Imports: add `from .xlsx_utils import ... style_output_sheets, apply_header_style, sample_data_font`.
2. At the end of `compare_workbooks`, after the `for title, ws_src, rows in sheets_to_write: _write_rows_sheet(...)` loop and before `wb_out.save(out_path)`:
```python
    apply_header_style(wb_out["Summary"])
    row_sheets = [wb_out[t] for t in wb_out.sheetnames if t != "Summary"]
    if row_sheets and sheets_to_write:
        style_output_sheets(
            row_sheets,
            width_by="header",
            data_font=sample_data_font(sheets_to_write[0][1], 1),
        )
```
3. `_write_summary` / `_write_rows_sheet` keep their own width and `Font(bold=True)` writes — `apply_header_style` and `style_output_sheets` run after and override the header row; `Summary`'s data-cell widths survive because it is not passed to `style_output_sheets`.

- [ ] **Step 4: Run**

Run: `python -m pytest tests/test_compare_versions.py -q` → PASS.

- [ ] **Step 5: Commit**

```bash
git add bigmedia/compare_versions.py tests/test_compare_versions.py
git commit -m "$(cat <<'EOF'
feat: style compare output headers and row sheets

Summary and every added/removed sheet get the black/white header; the
row sheets also get width-fit and a uniform data font. Summary keeps its
fixed column widths.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 6: Wire `fix_getty_split`

**Files:**
- Modify: `bigmedia/getty_split.py`
- Test: `tests/test_getty_split.py`

**Interfaces:**
- Consumes: `style_output_sheets`, `first_data_font` (Task 1).
- Produces: `fix_getty_split(old_path, new_path, out_path, name_column="Clip Name")` — unchanged signature. The Getty Videos / Getty Stills sheets get full styling; zebra row fills are preserved (styling sets font + header, not data fills).

- [ ] **Step 1: Write the failing test**

Append to `tests/test_getty_split.py` (reuse its existing fixtures):

```python
def test_getty_split_output_is_styled_zebra_kept(getty_split_pair, tmp_path):
    from openpyxl import load_workbook
    old_p, new_p = getty_split_pair
    out = tmp_path / "fixed.xlsx"
    fix_getty_split(str(old_p), str(new_p), str(out))
    wb = load_workbook(str(out))
    getty_sheets = [t for t in wb.sheetnames if t.lower().startswith("getty")]
    assert getty_sheets
    for t in getty_sheets:
        ws = wb[t]
        h = ws.cell(row=1, column=1)
        assert h.fill.fgColor.rgb == "FF000000"
        assert h.font.bold and h.font.color.rgb == "FFFFFFFF"
        assert 10 <= ws.column_dimensions["A"].width <= 60
        # zebra: rows 2 and 3 have different fills (data fills untouched)
        if ws.max_row >= 3:
            assert ws.cell(row=2, column=1).fill != ws.cell(row=3, column=1).fill
```

(Match the fixture names actually in `tests/test_getty_split.py`.)

- [ ] **Step 2: Run — verify it fails**

Run: `python -m pytest tests/test_getty_split.py -k styled -v` → FAIL.

- [ ] **Step 3: Implement**

In `bigmedia/getty_split.py`:
1. Imports: `from .xlsx_utils import find_column_any, copy_sheet_verbatim, style_output_sheets, first_data_font`.
2. In `fix_getty_split`, before `wb_out.save(out_path)`:
```python
    getty_sheets = [wb_out[t] for t in wb_out.sheetnames
                    if t in (videos_title, stills_title)]
    if getty_sheets:
        style_output_sheets(getty_sheets, width_by="index",
                            data_font=first_data_font(getty_sheets))
```

- [ ] **Step 4: Run**

Run: `python -m pytest tests/test_getty_split.py -q` → PASS.

- [ ] **Step 5: Commit**

```bash
git add bigmedia/getty_split.py tests/test_getty_split.py
git commit -m "$(cat <<'EOF'
feat: style fix-getty output sheets

Getty Videos / Getty Stills get the standard width-fit, uniform font and
black/white header. Zebra row striping is preserved.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 7: Remove the options from CLI and web form

**Files:**
- Modify: `bigmedia/cli.py`, `web/commands.py`
- Test: `tests/test_cli.py`, `tests/test_web_commands.py`, `tests/test_web_presets.py`, `tests/test_web_presets_ui.py`

**Interfaces:**
- Consumes: the trimmed `sort_workbook` / `group_duplicates_workbook` signatures (Tasks 2, 4).
- Produces: no `--autofit` / `--min-width` / `--max-width` / `--uniform-font` / `--uniform-header` CLI flags; no `autofit` / `min_width` / `max_width` / `uniform_font` fields on the `sort` web spec; no `autofit` / `min_width` / `max_width` / `uniform_font` / `uniform_header` on the `group` web spec.

- [ ] **Step 1: Update the tests**

- `tests/test_cli.py`: any test passing `--autofit` / `--min-width` etc. now expects a `SystemExit` (argparse error) — or just drop those assertions and keep a test that `sort` / `group` still run without them.
- `tests/test_web_commands.py`: update `test_sort_fields_match_cli_options` (and the group equivalent) to the trimmed field-name sets:
  - sort: `{"name_column", "categories", "skip_categories"}`
  - group: `{"name_column", "duration_column", "fps", "sheets"}`
- `tests/test_web_presets.py` / `tests/test_web_presets_ui.py`: check the 2 hits — if a test saves/loads a preset with `autofit` etc., either drop that key or assert the loader ignores unknown keys. Add if missing:
  ```python
  def test_preset_with_removed_key_loads_without_error(tmp_path):
      # a preset saved before the fields were removed still loads
      from web.presets import save_preset, get_preset, init_db
      db = tmp_path / "p.db"; init_db(db)
      save_preset(db, "sort", "old", {"name_column": "Clip Name", "autofit": True})
      assert get_preset(db, "sort", "old")["name_column"] == "Clip Name"
  ```

- [ ] **Step 2: Run — verify failures**

Run: `python -m pytest tests/test_cli.py tests/test_web_commands.py tests/test_web_presets.py tests/test_web_presets_ui.py -v`
Expected: the field-set and flag tests fail.

- [ ] **Step 3: Implement**

- `bigmedia/cli.py`: in the shared-options loop (around line 320) delete the `--autofit`, `--min-width`, `--max-width` lines; find and delete `--uniform-font` and `--uniform-header` (grep for them). In `cmd_sort` drop `autofit=`, `min_width=`, `max_width=`, `uniform_font=` from the `sort_workbook(...)` call; in `cmd_group` drop those plus `uniform_header=`. Remove the `AUTOFIT_MIN_WIDTH` / `AUTOFIT_MAX_WIDTH` import if unused.
- `web/commands.py`: remove the four `FieldSpec`s from `COMMANDS["sort"].fields` and the five from `COMMANDS["group"].fields`. Remove the `AUTOFIT_MAX_WIDTH` / `AUTOFIT_MIN_WIDTH` import and the `AUTOFIT_*` references if unused.

- [ ] **Step 4: Run — targeted then full**

Run: `python -m pytest tests/test_cli.py tests/test_web_commands.py tests/test_web_presets.py tests/test_web_presets_ui.py -q` → PASS.
Run: `python -m pytest -q` → **all green**.

- [ ] **Step 5: Commit**

```bash
git add bigmedia/cli.py web/commands.py tests/test_cli.py tests/test_web_commands.py tests/test_web_presets.py tests/test_web_presets_ui.py
git commit -m "$(cat <<'EOF'
refactor: drop the output-formatting flags and form fields

The autofit / min-width / max-width / uniform-font / uniform-header
options are gone from the CLI and the web forms -- the styling they
controlled is now always applied. Old presets carrying these keys still
load (unknown keys are ignored).

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 8: Docs + whole-suite verification

**Files:**
- Modify: `CLAUDE.md`
- Test: `tests/test_web_wizard_e2e.py` (or a new small e2e), full suite

**Interfaces:**
- Consumes: everything.

- [ ] **Step 1: Update `CLAUDE.md`**

In the "Conventions" bullet list, replace any mention of optional autofit / uniform styling with:

```markdown
- Output styling is standardized and always applied to generated sheets:
  column widths fit the wider of content or header (clamped 10-60), every
  data cell shares one font sampled from the source, and header rows are a
  solid black fill with white bold text. The "Worksheet" backup tab is
  exempt -- it stays a verbatim copy of the input. `fu-grid` and
  `getty-ids` keep their own bespoke layouts.
```

- [ ] **Step 2: Add an end-to-end styling check**

Append to `tests/test_web_wizard_e2e.py`:

```python
def test_run_sort_output_is_styled(monkeypatch):
    import io
    from openpyxl import load_workbook
    client = _client(monkeypatch)
    content = _master()  # existing helper in this file
    r = client.post("/commands/sort",
                    data={"name_column": "Media File", "categories": ["AP"], "categories_present": "1"},
                    files={"files": ("m.xlsx", content, "application/octet-stream")})
    assert r.status_code == 200
    wb = load_workbook(io.BytesIO(r.content))
    ap = wb["AP"]
    assert ap.cell(row=1, column=1).fill.fgColor.rgb == "FF000000"
    assert ap.cell(row=1, column=1).font.color.rgb == "FFFFFFFF"
    assert 10 <= ap.column_dimensions["A"].width <= 60
```

- [ ] **Step 3: Full suite**

Run: `python -m pytest -q`
Expected: all green.

- [ ] **Step 4: Manual smoke (curl)**

```bash
BIGMEDIA_WEB_PASSWORD_HASH=$(python -c "from web.auth import hash_password; print(hash_password('dev'))") \
BIGMEDIA_SESSION_SECRET=x python -m uvicorn web.main:app --port 8791 &
# login, POST a small xlsx to /commands/sort, open the result, confirm
# black headers + fitted widths; stop the server.
```
Put the result in the report.

- [ ] **Step 5: Commit**

```bash
git add CLAUDE.md tests/test_web_wizard_e2e.py
git commit -m "$(cat <<'EOF'
docs: standardized output styling; e2e styling check

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Self-Review

**1. Spec coverage:**

| Spec item | Task |
|---|---|
| `style_output_sheets`, `apply_header_style`, `sample_data_font`, `first_data_font`, `HEADER_FILL` | Task 1 |
| `measure_*` lose min/max/padding params | Task 1 |
| delete `dominant_header_fill` / `apply_uniform_header_fill` / `DEFAULT_HEADER_FILL_COLOR` | Task 1 |
| sort: drop params, style always (index) | Task 2 |
| dedupe: style always (index), no sig change | Task 3 |
| group: drop params, style always (header) | Task 4 |
| compare: Summary header + row sheets (header) | Task 5 |
| fix-getty: style Getty sheets (index), zebra kept | Task 6 |
| CLI flag removal | Task 7 |
| web form field removal | Task 7 |
| preset back-compat (unknown keys ignored) | Task 7 |
| CLAUDE.md conventions update | Task 8 |
| Worksheet / backup tabs verbatim | Task 1 guard + asserted in Tasks 2, 3 |
| fu-grid / getty-ids untouched | Not modified in any task (verified: no task touches `fu_grid.py` / `getty_ids.py`) |

No uncovered spec requirements.

**2. Placeholder scan:** Tasks reference real fixtures (`sample_master`, `sample_sorted_master`) from `conftest.py`; compare/getty-split fixture names are flagged "match the file's existing fixture" because those test files weren't fully read — the implementer confirms the exact name when writing the test. No "TBD"/"handle edge cases"/bare "write tests".

**3. Type consistency:**
- `style_output_sheets(sheets, *, width_by, data_font)` — same call shape in Tasks 2–6.
- `sample_data_font(ws, col_idx)` vs `first_data_font(worksheets)` — sort/dedupe/compare use the former, group/fix-getty use the latter; both defined in Task 1, both return a `Font`.
- `HEADER_FILL` fgColor `"FF000000"`, header font color `"FFFFFFFF"` — identical strings in Task 1 impl and every task's assertions.
- `width_by` values `"index"` / `"header"` — consistent.

---

## Execution Handoff

Plan complete. Two execution options:

1. **Subagent-Driven (recommended)** — fresh subagent per task, review between tasks.
2. **Inline Execution** — tasks in this session with checkpoints.
