# Sheet-aware analyze suggestions (group / fu-grid / getty-ids) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give `group`, `fu-grid`, and `getty-ids` header/column suggestions scoped to the sheet(s) they actually read (not Excel's last-active tab), plus real sheet-picker dropdowns/checklists populated from each upload's actual sheet names, for every sheet selection those three commands make.

**Architecture:** Extend `bigmedia/xlsx_utils.py`'s `analyze_files()` to compute per-sheet header data and per-sheet warnings alongside its existing active-sheet-based output (unchanged, still used by `sort`/`dedupe`). Add a `sheet_source` declaration to `FieldSpec` so `web/analyze.py` can resolve, from the current form values, which real sheet(s) a "headers" field should scope to, and a shared `match_sheet_name()` helper (also adopted by `getty_ids.py`, replacing its ad hoc equivalent) for the case-insensitive-plus-alias matching every command already does its own way. `group`'s free-text sheets field becomes a dynamically-populated checkbox list (same JSON `suggestions`/`sheets` contract Sort's category checkboxes already use), and `getty-ids`' sheet name fields become real dropdowns. No new HTTP endpoint or JSON top-level shape — this extends the existing `/commands/{slug}/analyze` contract from the previous analyze-wizard spec.

**Tech Stack:** Python 3, FastAPI, openpyxl, Jinja2, vanilla JS (no framework/build step), pytest, FastAPI `TestClient`.

**Spec:** `docs/superpowers/specs/2026-09-14-sheet-aware-analyze-suggestions-design.md` (follows up `docs/superpowers/specs/2026-09-10-analyze-suggest-wizard-design.md`)

## Global Constraints

- In scope: `group`, `fu-grid`, `getty-ids`. Out of scope: `sort`, `dedupe`, `compare`, `fix-getty` — do not touch `pair_command.html`, the `pair`/`fix_getty` upload modes, or `sort`/`dedupe`'s field wiring.
- `sort`/`dedupe` behavior must stay byte-for-byte unchanged: no `sheet_source` on their fields, so `run_analysis` must take the exact pre-existing active-sheet code path for them.
- Header/duration suggestions use the **union** of headers across every currently-selected sheet for a command, never the intersection.
- A stills-sheet (or any field with `allow_missing_sheet=True`) resolving to no match is a normal, silent state — never a warning, never a forced wrong guess.
- No new file I/O passes: `analyze_files` still opens each workbook exactly once, `read_only=True, data_only=True`.
- Full `pytest` suite must stay green after every task (444 passing at plan start).
- This repo has no JS test harness (no `package.json`/test runner) — JS changes are exercised indirectly through the Python-rendered template output and the JSON payload the JS consumes; verify JS behavior manually in a browser as the final task.

---

## Task 1: `match_sheet_name()` + sheet-scoped data in `analyze_files()`

**Files:**
- Modify: `bigmedia/xlsx_utils.py`
- Test: `tests/test_xlsx_utils.py`

**Interfaces:**
- Produces: `match_sheet_name(sheet_names: list[str], target: str, aliases: tuple[str, ...] = ()) -> str | None`
- Produces: `analyze_files(paths)` return dict gains three new keys alongside the existing `headers`/`sheets`/`warnings` (all unchanged in shape/content):
  - `headers_by_sheet: dict[str, list[str]]` — canonical (first-seen-casing) sheet name → union of that sheet's headers across every file that has it, first-seen order.
  - `sheet_warnings: dict[str, list[str]]` — canonical sheet name → warning strings (`"{file}: '{sheet}' sheet columns differ from {first_file}"` / `"{file}: '{sheet}' sheet not found"`), for every sheet seen in at least one file.
  - `unreadable_warnings: list[str]` — just the `"{file}: could not read (...)"` subset of `warnings`.
- Consumes: nothing new (still `openpyxl.load_workbook`, `pathlib.Path`, already imported).

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_xlsx_utils.py` (uses the existing `_make(path, sheets)` helper already in that file):

```python
from bigmedia.xlsx_utils import match_sheet_name  # add to the existing import block


def test_match_sheet_name_case_insensitive():
    assert match_sheet_name(["Getty Videos", "COST"], "getty videos") == "Getty Videos"


def test_match_sheet_name_strips_whitespace():
    assert match_sheet_name([" Getty Videos "], "Getty Videos") == " Getty Videos "


def test_match_sheet_name_falls_back_to_alias():
    assert match_sheet_name(["Getty pics"], "Getty Stills", aliases=("Getty pics",)) == "Getty pics"


def test_match_sheet_name_prefers_target_over_alias():
    names = ["Getty Stills", "Getty pics"]
    assert match_sheet_name(names, "Getty Stills", aliases=("Getty pics",)) == "Getty Stills"


def test_match_sheet_name_no_match_returns_none():
    assert match_sheet_name(["COST", "worksheet"], "Getty Videos") is None


def test_headers_by_sheet_per_sheet_not_just_active(tmp_path):
    # Active sheet is "COST" (first sheet created); "Getty Videos" is a
    # different, non-active sheet -- headers_by_sheet must still cover it.
    a = _make(tmp_path / "a.xlsx", {
        "COST": [["EP1"]],
        "Getty Videos": [["Track", "Clip Name", "Seconds"], ["V1", "x.mov", 6]],
    })
    result = analyze_files([a])
    assert result["headers"] == ["EP1"]  # unchanged: active-sheet only
    assert result["headers_by_sheet"]["COST"] == ["EP1"]
    assert result["headers_by_sheet"]["Getty Videos"] == ["Track", "Clip Name", "Seconds"]


def test_headers_by_sheet_unions_across_files_same_sheet_name(tmp_path):
    a = _make(tmp_path / "a.xlsx", {"Getty Videos": [["Clip Name", "Seconds"]]})
    b = _make(tmp_path / "b.xlsx", {"Getty Videos": [["Clip Name", "Notes"]]})
    result = analyze_files([a, b])
    assert result["headers_by_sheet"]["Getty Videos"] == ["Clip Name", "Seconds", "Notes"]


def test_headers_by_sheet_keyed_by_first_seen_casing(tmp_path):
    a = _make(tmp_path / "a.xlsx", {"Getty Videos": [["Clip Name"]]})
    b = _make(tmp_path / "b.xlsx", {"getty videos": [["Clip Name"]]})
    result = analyze_files([a, b])
    assert list(result["headers_by_sheet"].keys()) == ["Getty Videos"]


def test_sheet_warnings_disagreement_scoped_to_that_sheet(tmp_path):
    a = _make(tmp_path / "ep1.xlsx", {"Getty Videos": [["Clip Name", "Seconds"]], "COST": [["EP1"]]})
    b = _make(tmp_path / "ep2.xlsx", {"Getty Videos": [["Name"]], "COST": [["EP2"]]})
    result = analyze_files([a, b])
    # Active-sheet (COST) headers agree in shape ("EP1" vs "EP2" are both
    # single-header sheets) -> no generic warning; the Getty Videos
    # disagreement only shows up in sheet_warnings.
    assert not result["warnings"]
    assert any("ep2.xlsx" in w and "Getty Videos" in w and "differ" in w
               for w in result["sheet_warnings"]["Getty Videos"])


def test_sheet_warnings_not_found_lists_the_missing_file(tmp_path):
    a = _make(tmp_path / "ep1.xlsx", {"Getty Videos": [["Clip Name"]], "Getty Stills": [["Clip Name"]]})
    b = _make(tmp_path / "ep2.xlsx", {"Getty Videos": [["Clip Name"]]})  # no stills sheet
    result = analyze_files([a, b])
    assert any("ep2.xlsx" in w and "Getty Stills" in w and "not found" in w
               for w in result["sheet_warnings"]["Getty Stills"])
    assert "Getty Videos" not in "".join(result["sheet_warnings"].get("Getty Videos", []))


def test_unreadable_warnings_is_the_could_not_read_subset(tmp_path):
    good = _make(tmp_path / "good.xlsx", {"S": [["Clip Name"], ["x"]]})
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a real xlsx")
    result = analyze_files([str(bad), good])
    assert len(result["unreadable_warnings"]) == 1
    assert "bad.xlsx" in result["unreadable_warnings"][0]
    assert "could not read" in result["unreadable_warnings"][0]
    assert result["unreadable_warnings"] == result["warnings"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_xlsx_utils.py -k "match_sheet_name or headers_by_sheet or sheet_warnings or unreadable_warnings" -v`
Expected: FAIL — `match_sheet_name` doesn't exist (ImportError) and the new dict keys are absent (`KeyError`).

- [ ] **Step 3: Implement `match_sheet_name()` and extend `analyze_files()`**

In `bigmedia/xlsx_utils.py`, add `match_sheet_name` right before `analyze_files` (it has no dependency on anything else in the file):

```python
def match_sheet_name(sheet_names, target, aliases=()):
    """Case-insensitive match of `target` against `sheet_names`, falling
    back through `aliases` (alternate names to try, in order) when `target`
    itself doesn't match anything. Returns the real sheet name (original
    casing, from `sheet_names`) or None. Centralizes the
    match-case-insensitively-then-try-aliases pattern every command that
    hunts for a named sheet (Getty Videos, Getty Stills, 3rd parties, ...)
    already implements ad hoc."""
    candidates = [target] + [a for a in aliases if a.strip().lower() != target.strip().lower()]
    lookup = {name.strip().lower(): name for name in sheet_names}
    for candidate in candidates:
        match = lookup.get(candidate.strip().lower())
        if match:
            return match
    return None
```

Replace the whole existing `analyze_files` function with:

```python
def analyze_files(paths):
    """Read-only inspection of one or more clip-list workbooks, for the web
    UI's analyze-and-suggest step. Returns the union of the active sheets'
    header names, every sheet's name and (header-excluded) row count summed
    across files, and warnings for files that won't open or whose columns
    disagree with the first readable file. Never raises for a bad file --
    it lands in "warnings" and the others are still processed.

    Also returns per-sheet data for commands that read a specific named
    sheet rather than "whatever tab was active when the file was last
    saved" (group/fu-grid/getty-ids): `headers_by_sheet` (union of headers
    per sheet, across every file that has that sheet), `sheet_warnings`
    (per-sheet disagreement/missing-file warnings), and
    `unreadable_warnings` (just the "could not read" subset of `warnings`,
    since those are relevant regardless of which sheet a command cares
    about).

    Header values are stringified (``str(value).strip()``) so a stray
    ``datetime``/number cell in row 1 still yields a JSON-serializable
    header; whitespace-only cells are still dropped."""
    headers = []
    seen_headers = set()
    first_header_set = None
    first_name = None
    sheet_rows = {}
    sheet_order = []
    warnings = []
    unreadable_warnings = []

    file_names = []
    sheet_key_by_lower = {}          # sheet title lowercased -> canonical (first-seen) title
    headers_by_sheet = {}            # canonical title -> [headers...] (union, first-seen order)
    seen_headers_by_sheet = {}       # canonical title -> set(headers already added)
    sheet_occurrences = {}           # canonical title -> [(file_name, set(headers)), ...]

    for path in paths:
        name = Path(path).name
        wb = None
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            file_names.append(name)

            active = wb.active
            file_headers = [
                str(c.value).strip() for c in next(active.iter_rows(min_row=1, max_row=1), [])
                if c.value is not None and str(c.value).strip() != ""
            ]
            for h in file_headers:
                if h not in seen_headers:
                    seen_headers.add(h)
                    headers.append(h)

            if first_header_set is None:
                first_header_set = set(file_headers)
                first_name = name
            elif set(file_headers) != first_header_set:
                warnings.append(f"{name}: columns differ from {first_name}")

            for title in wb.sheetnames:
                ws = wb[title]
                canonical = sheet_key_by_lower.setdefault(title.strip().lower(), title)

                sheet_headers = [
                    str(c.value).strip() for c in next(ws.iter_rows(min_row=1, max_row=1), [])
                    if c.value is not None and str(c.value).strip() != ""
                ]
                bucket = headers_by_sheet.setdefault(canonical, [])
                seen_for_sheet = seen_headers_by_sheet.setdefault(canonical, set())
                for h in sheet_headers:
                    if h not in seen_for_sheet:
                        seen_for_sheet.add(h)
                        bucket.append(h)
                sheet_occurrences.setdefault(canonical, []).append((name, set(sheet_headers)))

                rows = ws.max_row
                if rows is None:
                    rows = sum(1 for _ in ws.iter_rows())
                count = max(rows - 1, 0)
                if title not in sheet_rows:
                    sheet_rows[title] = 0
                    sheet_order.append(title)
                sheet_rows[title] += count

        except Exception as exc:  # openpyxl raises several unrelated types
            msg = f"{name}: could not read ({exc})"
            warnings.append(msg)
            unreadable_warnings.append(msg)
        finally:
            if wb is not None:
                wb.close()

    sheet_warnings = {}
    for canonical, occurrences in sheet_occurrences.items():
        msgs = []
        first_file, first_headers = occurrences[0]
        present_files = {fn for fn, _ in occurrences}
        for fn, hset in occurrences[1:]:
            if hset != first_headers:
                msgs.append(f"{fn}: '{canonical}' sheet columns differ from {first_file}")
        for fn in file_names:
            if fn not in present_files:
                msgs.append(f"{fn}: '{canonical}' sheet not found")
        sheet_warnings[canonical] = msgs

    return {
        "headers": headers,
        "headers_by_sheet": headers_by_sheet,
        "sheet_warnings": sheet_warnings,
        "sheets": [{"name": t, "rows": sheet_rows[t]} for t in sheet_order],
        "warnings": warnings,
        "unreadable_warnings": unreadable_warnings,
    }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_xlsx_utils.py -v`
Expected: PASS — all new tests, and every pre-existing `test_xlsx_utils.py` test (the old keys/behavior are untouched).

- [ ] **Step 5: Commit**

```bash
git add bigmedia/xlsx_utils.py tests/test_xlsx_utils.py
git commit -m "feat: sheet-scoped header/warning data in analyze_files, add match_sheet_name"
```

---

## Task 2: Rewrite `getty_ids.py`'s sheet matching on `match_sheet_name`

Pure refactor — behavior-preserving, no test changes needed (the existing `tests/test_getty_ids.py` suite is the regression net).

**Files:**
- Modify: `bigmedia/getty_ids.py`

**Interfaces:**
- Consumes: `match_sheet_name` from Task 1.
- Produces: `_find_sheet(wb, sheet_name)` and `_find_sheet_any(wb, sheet_name, aliases=())` keep their exact existing signatures and return values (real sheet title or `None`) — callers (`_read_getty_ids`) are untouched.

- [ ] **Step 1: Run the existing test suite to capture the pre-refactor baseline**

Run: `pytest tests/test_getty_ids.py -v`
Expected: PASS (all currently passing, per the session's earlier work on this file).

- [ ] **Step 2: Rewrite `_find_sheet`/`_find_sheet_any` in terms of `match_sheet_name`**

In `bigmedia/getty_ids.py`, change the import line:

```python
from .xlsx_utils import find_column, iter_xlsx_files, match_sheet_name
```

Replace:

```python
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
```

with:

```python
def _find_sheet(wb, sheet_name):
    """Sheet titles vary in case across real delivery/sort output (e.g.
    "Getty Videos" vs "Getty videos") — match case-insensitively, same
    convention as group_duplicates.py."""
    return match_sheet_name(wb.sheetnames, sheet_name)


# Real delivery files sometimes use an entirely different word for the
# stills sheet, not just different casing (e.g. "Getty pics" instead of
# "Getty Stills") — a plain case-insensitive match on the caller's chosen
# name won't catch that, so fall back through known aliases same as
# NAME_COLUMN_ALIASES above.
_STILLS_SHEET_ALIASES = ("Getty Stills", "Getty pics")


def _find_sheet_any(wb, sheet_name, aliases=()):
    return match_sheet_name(wb.sheetnames, sheet_name, aliases=aliases)
```

- [ ] **Step 3: Run the test suite again to confirm the refactor is behavior-preserving**

Run: `pytest tests/test_getty_ids.py -v`
Expected: PASS — identical results to Step 1, no test changes.

- [ ] **Step 4: Commit**

```bash
git add bigmedia/getty_ids.py
git commit -m "refactor: getty_ids sheet matching now shares xlsx_utils.match_sheet_name"
```

---

## Task 3: Fix `group_duplicates_workbook`'s empty-vs-omitted `sheets` handling

An explicit empty list must mean "group nothing," not silently fall back to `DEFAULT_SHEETS` — needed before Task 6 wires the web checklist's "everything unchecked" case to an explicit `[]`.

**Files:**
- Modify: `bigmedia/group_duplicates.py`
- Test: `tests/test_group_duplicates.py`

**Interfaces:**
- Produces: `group_duplicates_workbook(src_path, out_path, sheets=None, ...)` — `sheets=None` still means "use `DEFAULT_SHEETS`"; `sheets=[]` now means "group no sheets, pass every sheet through verbatim." (Previously `sheets=[]` was indistinguishable from `None` because of `sheets or DEFAULT_SHEETS`.)

- [ ] **Step 1: Write the failing test**

Add to `tests/test_group_duplicates.py`:

```python
def test_empty_sheets_list_groups_nothing(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out_path), sheets=[])

    wb = load_workbook(out_path)
    ws = wb["Getty Videos"]
    # No grouping happened: no Total Duration/Seconds columns were inserted,
    # unlike test_group_duplicates_reorders_and_sums_getty_videos above.
    header_row = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    assert "Total Duration" not in header_row
    assert "Seconds" not in header_row
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/test_group_duplicates.py::test_empty_sheets_list_groups_nothing -v`
Expected: FAIL — `sheets=[]` currently falls back to `DEFAULT_SHEETS` via `sheets or DEFAULT_SHEETS`, so "Getty Videos" still gets grouped and the assertion fails.

- [ ] **Step 3: Fix the default handling**

In `bigmedia/group_duplicates.py`, in `group_duplicates_workbook`, change:

```python
    sheets = sheets or DEFAULT_SHEETS
```

to:

```python
    sheets = DEFAULT_SHEETS if sheets is None else sheets
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_group_duplicates.py -v`
Expected: PASS — the new test, and every existing one (which all call with `sheets=None`/omitted, unaffected by the `is None` change).

- [ ] **Step 5: Commit**

```bash
git add bigmedia/group_duplicates.py tests/test_group_duplicates.py
git commit -m "fix: group_duplicates_workbook treats an explicit empty sheets list as 'group nothing'"
```

---

## Task 4: `FieldSpec.sheet_source`/`allow_missing_sheet` + per-command field wiring

**Files:**
- Modify: `web/commands.py`
- Test: `tests/test_web_commands.py`

**Interfaces:**
- Produces: `FieldSpec` gains `sheet_source: list = None` and `allow_missing_sheet: bool = False`.
- Produces: `COMMANDS["group"].fields` gets a `sheets` field of `type="sheet_checklist"`, plus local `name_column`/`duration_column` with `sheet_source=["sheets"]`.
- Produces: `COMMANDS["fu-grid"].fields` gets local `name_column`/`duration_column` with `sheet_source=["sheet"]`.
- Produces: `COMMANDS["getty-ids"].fields`: `sheet_name`/`stills_sheet_name` become `options_source="sheets"` (stills: `allow_missing_sheet=True`); `name_column`/new `seconds_column` (now `options_source="headers"`) get `sheet_source=["sheet_name", "stills_sheet_name"]`.
- Consumes: nothing new from other tasks (pure data/config change) — Task 5 (`web/analyze.py`) is what actually reads `sheet_source`/`allow_missing_sheet`.

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_web_commands.py`:

```python
def test_getty_ids_sheet_fields_are_dropdowns_not_text():
    fields = {f.name: f for f in COMMANDS["getty-ids"].fields}
    assert fields["sheet_name"].options_source == "sheets"
    assert fields["sheet_name"].allow_missing_sheet is False
    assert fields["stills_sheet_name"].options_source == "sheets"
    assert fields["stills_sheet_name"].allow_missing_sheet is True


def test_getty_ids_headers_fields_scope_to_both_sheets():
    fields = {f.name: f for f in COMMANDS["getty-ids"].fields}
    assert fields["name_column"].sheet_source == ["sheet_name", "stills_sheet_name"]
    assert fields["seconds_column"].options_source == "headers"
    assert fields["seconds_column"].sheet_source == ["sheet_name", "stills_sheet_name"]


def test_fu_grid_headers_fields_scope_to_sheet_field():
    fields = {f.name: f for f in COMMANDS["fu-grid"].fields}
    assert fields["name_column"].sheet_source == ["sheet"]
    assert fields["duration_column"].sheet_source == ["sheet"]


def test_group_sheets_field_is_a_checklist_scoping_name_and_duration():
    fields = {f.name: f for f in COMMANDS["group"].fields}
    assert fields["sheets"].type == "sheet_checklist"
    assert fields["sheets"].options_source == "sheets"
    assert fields["name_column"].sheet_source == ["sheets"]
    assert fields["duration_column"].sheet_source == ["sheets"]


def test_sort_and_dedupe_fields_declare_no_sheet_source():
    # Explicitly unchanged: no sheet_source anywhere for the two commands
    # this feature does not touch.
    for slug in ("sort", "dedupe"):
        for f in COMMANDS[slug].fields:
            assert not f.sheet_source
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_web_commands.py -k "sheet_fields or headers_fields_scope or checklist_scoping or declare_no_sheet_source" -v`
Expected: FAIL — `AttributeError`/`AssertionError`, `sheet_source`/`allow_missing_sheet` don't exist yet and the fields are still plain text/shared.

- [ ] **Step 3: Update `FieldSpec` and the three commands' field lists**

In `web/commands.py`, change the `FieldSpec` dataclass:

```python
@dataclass
class FieldSpec:
    name: str
    label: str
    type: str  # "text" | "number" | "checkbox" | "list" | "sheet_checklist"
    default: Any = None
    required: bool = False
    options_source: str = None  # None | "headers" | "sheets"
    sheet_source: list = None  # field name(s) whose current sheet
                                # selection scopes a "headers" field's
                                # options (group/fu-grid/getty-ids only)
    allow_missing_sheet: bool = False  # "sheets"-sourced fields only:
                                        # True allows an empty "(none)"
                                        # selection (getty-ids' stills sheet)
```

Replace the `"group"` entry:

```python
    "group": CommandSpec(
        slug="group", title="Group", upload_mode="batch", func=group_duplicates_workbook,
        output_suffix="grouped",
        fields=[
            FieldSpec("name_column", "Filename column", "text", "Clip Name",
                      options_source="headers", sheet_source=["sheets"]),
            FieldSpec("duration_column", "Duration column", "text", "Clip Duration",
                      options_source="headers", sheet_source=["sheets"]),
            _FPS,
            FieldSpec("sheets", "Sheets to group", "sheet_checklist", options_source="sheets"),
        ],
    ),
```

Replace the `"fu-grid"` entry:

```python
    "fu-grid": CommandSpec(
        slug="fu-grid", title="FU Grid", upload_mode="batch", func=fu_grid_workbook,
        output_suffix="fu_grid",
        fields=[
            FieldSpec("sheet", "Source sheet", "text", "3rd parties", options_source="sheets"),
            FieldSpec("name_column", "Filename column", "text", "Clip Name",
                      options_source="headers", sheet_source=["sheet"]),
            FieldSpec("duration_column", "Duration column", "text", "Clip Duration",
                      options_source="headers", sheet_source=["sheet"]),
            _FPS,
        ],
    ),
```

Replace the `"getty-ids"` entry's field list:

```python
    "getty-ids": CommandSpec(
        slug="getty-ids", title="Getty IDs", upload_mode="combine", func=build_getty_id_report,
        output_suffix="getty_ids",
        fields=[
            FieldSpec("project_name", "Project name", "text", required=True),
            FieldSpec("sheet_name", "Video sheet name", "text", "Getty Videos",
                      options_source="sheets"),
            FieldSpec("stills_sheet_name", "Stills sheet name", "text", "Getty Stills",
                      options_source="sheets", allow_missing_sheet=True),
            FieldSpec("name_column", "Filename column", "text", "Clip Name",
                      options_source="headers", sheet_source=["sheet_name", "stills_sheet_name"]),
            FieldSpec("seconds_column", "Seconds column", "text", "Seconds",
                      options_source="headers", sheet_source=["sheet_name", "stills_sheet_name"]),
            FieldSpec("min_seconds", "Min seconds (videos only)", "number", 5),
            FieldSpec("max_seconds", "Max seconds (videos only, blank = no limit)", "number"),
            FieldSpec("include_video", "Include video clips", "checkbox", True),
            FieldSpec("include_stills", "Include stills", "checkbox", True),
        ],
    ),
```

`sort`, `dedupe`, `compare`, `fix-getty` are untouched — they keep using the shared `_NAME_COLUMN`/`_DURATION_COLUMN` constants exactly as before.

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_web_commands.py -v`
Expected: PASS — new tests, and every pre-existing `test_web_commands.py` test (they check `upload_mode`, `required`, `options_source == "headers"` for `name_column`, none of which this change breaks).

- [ ] **Step 5: Commit**

```bash
git add web/commands.py tests/test_web_commands.py
git commit -m "feat: FieldSpec.sheet_source/allow_missing_sheet; wire group/fu-grid/getty-ids sheet fields"
```

---

## Task 5: `web/analyze.py` — resolve `sheet_source` into scoped headers/warnings/suggestions

**Files:**
- Modify: `web/analyze.py`
- Test: `tests/test_web_analyze.py`

**Interfaces:**
- Consumes: `analyze_files()`'s new `headers_by_sheet`/`sheet_warnings`/`unreadable_warnings` (Task 1), `match_sheet_name` (Task 1), `FieldSpec.sheet_source`/`allow_missing_sheet` (Task 4), `group_duplicates.DEFAULT_SHEETS`/`SHEET_ALIASES`.
- Produces: `run_analysis(spec, paths, form)` — same signature as before, and the same `headers`/`sheets`/`suggestions`/`annotations`/`warnings` keys, plus one new key `headers_by_sheet` (passed through from `analyze_files()` verbatim, per the spec's JSON contract) for a future consumer to inspect per-sheet headers directly; for a command with at least one `sheet_source`-declaring field, `headers`/`warnings` are now sheet-scoped instead of active-sheet-based, and `suggestions` gains an entry for every `options_source == "sheets"` field that has a match.

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_web_analyze.py` (reuses the file's existing `_sheet(path, headers, names)` helper, which writes to `wb.active` — extend the test file with a small multi-sheet helper alongside it):

```python
def _multi_sheet(path, sheets):
    """sheets: dict of {title: [header_row, *data_rows]}, first key becomes
    the active sheet (matches openpyxl's default: first-created = active)."""
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title=title)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return str(path)


def test_getty_ids_headers_scoped_to_sheet_name_not_active_sheet(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {
        "COST": [["EP1"]],
        "Getty Videos": [["Clip Name", "Seconds"], ["x.mov", 6]],
    })
    result = run_analysis(COMMANDS["getty-ids"], [path], {"sheet_name": "Getty Videos", "stills_sheet_name": "Getty Stills"})
    assert result["headers"] == ["Clip Name", "Seconds"]
    assert result["suggestions"]["name_column"] == "Clip Name"


def test_getty_ids_stills_sheet_missing_everywhere_suggests_none_no_warning(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {"Getty Videos": [["Clip Name"]]})
    result = run_analysis(COMMANDS["getty-ids"], [path], {"sheet_name": "Getty Videos", "stills_sheet_name": "Getty Stills"})
    assert result["suggestions"]["stills_sheet_name"] == ""
    assert not result["warnings"]


def test_getty_ids_missing_required_video_sheet_warns(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {"COST": [["EP1"]]})
    result = run_analysis(COMMANDS["getty-ids"], [path], {"sheet_name": "Getty Videos", "stills_sheet_name": "Getty Stills"})
    assert any("Getty Videos" in w and "not found" in w for w in result["warnings"])


def test_getty_ids_sheet_name_suggests_case_insensitive_match(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {"getty videos": [["Clip Name"]]})
    result = run_analysis(COMMANDS["getty-ids"], [path], {})
    assert result["suggestions"]["sheet_name"] == "getty videos"


def test_getty_ids_stills_sheet_suggests_via_alias(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {"Getty Videos": [["Clip Name"]], "Getty pics": [["Clip Name"]]})
    result = run_analysis(COMMANDS["getty-ids"], [path], {})
    assert result["suggestions"]["stills_sheet_name"] == "Getty pics"


def test_fu_grid_headers_scoped_to_selected_sheet(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {
        "3rd parties": [["Clip Name", "Clip Duration"], ["x.mov", 5]],
        "COST": [["EP1"]],
    })
    result = run_analysis(COMMANDS["fu-grid"], [path], {"sheet": "3rd parties"})
    assert result["headers"] == ["Clip Name", "Clip Duration"]


def test_fu_grid_sheet_field_suggests_default_when_present(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {"3rd parties": [["Clip Name"]]})
    result = run_analysis(COMMANDS["fu-grid"], [path], {})
    assert result["suggestions"]["sheet"] == "3rd parties"


def test_group_headers_scoped_to_checked_sheets_union(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {
        "AP": [["Clip Name"]],
        "Getty Videos": [["Clip Name", "Clip Duration"]],
        "COST": [["EP1"]],
    })
    result = run_analysis(COMMANDS["group"], [path], {"sheets": ["AP", "Getty Videos"]})
    assert set(result["headers"]) == {"Clip Name", "Clip Duration"}


def test_group_sheets_checklist_suggests_default_sheets_present(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {
        "AP": [["Clip Name"]],
        "Getty Videos": [["Clip Name"]],
        "COST": [["EP1"]],
    })
    result = run_analysis(COMMANDS["group"], [path], {})
    assert set(result["suggestions"]["sheets"]) == {"AP", "Getty Videos"}
    assert "COST" not in result["suggestions"]["sheets"]


def test_group_sheets_checklist_matches_stills_alias(tmp_path):
    path = _multi_sheet(tmp_path / "ep1.xlsx", {"Getty pics": [["Clip Name"]]})
    result = run_analysis(COMMANDS["group"], [path], {})
    assert result["suggestions"]["sheets"] == ["Getty pics"]


def test_sort_dedupe_headers_still_active_sheet_based(tmp_path):
    # No sheet_source anywhere for these two -> untouched code path.
    path = _multi_sheet(tmp_path / "ep1.xlsx", {
        "COST": [["EP1"]],
        "Getty Videos": [["Clip Name"]],
    })
    result = run_analysis(COMMANDS["dedupe"], [path], {})
    assert result["headers"] == ["EP1"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_web_analyze.py -v`
Expected: FAIL on all the new tests (headers still active-sheet-based, no `sheet_name`/`sheets`/`sheet` suggestions computed) — the pre-existing tests in this file still PASS (nothing touched yet).

- [ ] **Step 3: Implement the sheet-scoping logic**

Replace the full contents of `web/analyze.py` with:

```python
import inspect

from bigmedia.group_duplicates import DEFAULT_SHEETS, DURATION_COLUMN_ALIASES, SHEET_ALIASES
from bigmedia.sort_workbook import NAME_COLUMN_ALIASES
from bigmedia.xlsx_utils import analyze_files, match_sheet_name
from web.files import parse_field

# Which suggestion field each alias tuple feeds. Reused from the business
# modules so the web layer and the workbook code agree on what counts as a
# name/duration header.
_COLUMN_SUGGESTIONS = (
    ("name_column", NAME_COLUMN_ALIASES),
    ("duration_column", DURATION_COLUMN_ALIASES),
)


def _form_value(form, field):
    """One value for `field` out of a submitted form, tolerating both a
    Starlette FormData (has getlist) and a plain dict (tests). A
    "sheet_checklist" is multi-value the same way a "list" field is."""
    if field.type in ("list", "sheet_checklist") and hasattr(form, "getlist"):
        return form.getlist(field.name)
    return form.get(field.name)


def _resolve_sheet_field_values(form, field):
    """Current value(s) of a sheet-selecting field (a "sheets" dropdown or
    a "sheet_checklist"), as a list, blanks dropped. A plain dict test form
    can supply either a single string or a list directly."""
    raw = _form_value(form, field)
    values = raw if isinstance(raw, list) else [raw]
    return [v for v in values if v]


def _sheet_source_result(spec, form, generic):
    """For a command whose "headers" fields declare sheet_source, resolve
    those fields' current sheet selection(s) against the real sheets seen
    across the uploaded files, and return the sheet-scoped headers list
    plus warnings limited to those sheets (union across every resolved
    sheet, never intersection). Returns None for a command with no
    sheet_source fields at all (sort/dedupe), so the caller keeps the
    generic active-sheet-based headers/warnings untouched."""
    fields_by_name = {f.name: f for f in spec.fields}
    headers_fields = [f for f in spec.fields if f.sheet_source]
    if not headers_fields:
        return None

    real_sheet_names = list(generic["headers_by_sheet"].keys())
    resolved = {}  # typed/selected sheet name -> (matched real name or None, allow_missing)
    for hf in headers_fields:
        for source_name in hf.sheet_source:
            source_field = fields_by_name[source_name]
            for target in _resolve_sheet_field_values(form, source_field):
                if target in resolved:
                    continue
                aliases = SHEET_ALIASES.get(target.strip().lower(), ())
                matched = match_sheet_name(real_sheet_names, target, aliases=aliases)
                resolved[target] = (matched, source_field.allow_missing_sheet)

    headers = []
    seen = set()
    for matched, _allow_missing in resolved.values():
        if not matched:
            continue
        for h in generic["headers_by_sheet"].get(matched, []):
            if h not in seen:
                seen.add(h)
                headers.append(h)

    warnings = list(generic["unreadable_warnings"])
    for target, (matched, allow_missing) in resolved.items():
        if matched is None:
            if not allow_missing:
                warnings.append(f"'{target}' sheet not found in any uploaded file")
            continue
        for msg in generic["sheet_warnings"].get(matched, []):
            if allow_missing and msg.endswith("sheet not found"):
                continue
            warnings.append(msg)

    return {"headers": headers, "warnings": warnings}


def _suggest_sheet_value(field, generic):
    """Suggested value(s) for an options_source="sheets" field: a single
    matched real sheet name for a plain dropdown, or a list of matches for
    a "sheet_checklist". Returns None when there's nothing to suggest (the
    template/JS leave the field at its current value/empty)."""
    real_sheet_names = list(generic["headers_by_sheet"].keys())

    if field.type == "sheet_checklist":
        matches = []
        for wanted in DEFAULT_SHEETS:
            aliases = SHEET_ALIASES.get(wanted.lower(), ())
            matched = match_sheet_name(real_sheet_names, wanted, aliases=aliases)
            if matched and matched not in matches:
                matches.append(matched)
        return matches or None

    target = field.default or ""
    if not target:
        return None
    aliases = SHEET_ALIASES.get(target.strip().lower(), ())
    matched = match_sheet_name(real_sheet_names, target, aliases=aliases)
    if matched:
        return matched
    return "" if field.allow_missing_sheet else None


def _analyzer_kwargs(spec, form):
    """The subset of the user's current field values that the analyzer
    actually declares as parameters, typed via parse_field.

    Guards:
    - only real field names are considered (an analyzer param named `files`
      can't grab an UploadFile);
    - POSITIONAL_ONLY / *args / **kwargs params are excluded;
    - empty values (None, "", [], unchecked checkbox) are dropped so the
      analyzer keeps its own defaults.
    """
    fields_by_name = {f.name: f for f in spec.fields}

    kwargs = {}
    for pname, param in inspect.signature(spec.analyze).parameters.items():
        if pname == "paths" or pname not in fields_by_name:
            continue
        if param.kind in (param.VAR_POSITIONAL, param.VAR_KEYWORD, param.POSITIONAL_ONLY):
            continue
        field = fields_by_name[pname]
        value = parse_field(field, _form_value(form, field))
        if value is None or value == "" or value == [] or value is False:
            continue
        kwargs[pname] = value
    return kwargs


def run_analysis(spec, paths, form):
    generic = analyze_files(paths)

    sheet_scoped = _sheet_source_result(spec, form, generic)
    headers = sheet_scoped["headers"] if sheet_scoped is not None else generic["headers"]
    warnings = sheet_scoped["warnings"] if sheet_scoped is not None else list(generic["warnings"])

    result = {
        "headers": headers,
        "headers_by_sheet": generic["headers_by_sheet"],
        "sheets": generic["sheets"],
        "suggestions": {},
        "annotations": {},
        "warnings": warnings,
    }

    # Generic name/duration column guess from known header aliases (spec
    # decision 4). setdefault + running before the per-command analyzer so a
    # command-specific analyzer's suggestions.update() can still override.
    for field, aliases in _COLUMN_SUGGESTIONS:
        match = next((a for a in aliases if a in headers), None)
        if match:
            result["suggestions"].setdefault(field, match)

    # Sheet-picker suggestions (getty-ids' video/stills sheet, fu-grid's
    # source sheet, group's checklist) -- same shape for every command that
    # has one, so computed generically here instead of via a per-command
    # spec.analyze hook.
    for field in spec.fields:
        if field.options_source == "sheets":
            suggestion = _suggest_sheet_value(field, generic)
            if suggestion is not None:
                result["suggestions"].setdefault(field.name, suggestion)

    if spec.analyze:
        extra = spec.analyze(paths, **_analyzer_kwargs(spec, form))
        result["suggestions"].update(extra.get("suggestions", {}))
        result["annotations"].update(extra.get("annotations", {}))
        result["warnings"].extend(extra.get("warnings", []))
    return result
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_web_analyze.py -v`
Expected: PASS — all new tests, and every pre-existing test in this file (in particular `test_generic_only_command_returns_headers_and_sheets` and the four `sort`-analyzer tests, which exercise `dedupe`/`sort` and must be byte-identical).

- [ ] **Step 5: Commit**

```bash
git add web/analyze.py tests/test_web_analyze.py
git commit -m "feat: web/analyze.py resolves sheet_source into scoped headers/warnings/suggestions"
```

---

## Task 6: Multi-value parsing for `sheet_checklist` + group's `sheets_present` marker

**Files:**
- Modify: `web/files.py`, `web/main.py`, `bigmedia/group_duplicates.py` (import only — `DEFAULT_SHEETS` already there from Task 3)
- Test: `tests/test_web_routes_group.py` (new)

**Interfaces:**
- Produces: `parse_field(field, raw)` in `web/files.py` treats `field.type == "sheet_checklist"` exactly like `"list"`.
- Produces: `_build_kwargs` in `web/main.py` uses `form.getlist(f.name)` for `f.type in ("list", "sheet_checklist")`.
- Produces: `command_submit` in `web/main.py`: when `form.get("sheets_present")` is truthy and `kwargs.get("sheets")` is falsy, sets `kwargs["sheets"] = []` (group only) — same pattern as the existing `categories_present` handling for sort.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_web_routes_group.py`:

```python
import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from web.auth import hash_password
from web.main import app


def _logged_in_client(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    client = TestClient(app)
    client.post("/login", data={"password": "pw"})
    return client


def _sorted_workbook_bytes():
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Getty Videos")
    ws.append(["Clip Name", "Clip Duration"])
    ws.append(["GettyImages-12345.mov", "00:00:00:05"])
    ap = wb.create_sheet("AP")
    ap.append(["Clip Name", "Clip Duration"])
    ap.append(["BM1234_x.mxf", "00:00:00:05"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_group_sheets_checkboxes_group_only_checked_sheets(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("ep1.xlsx", _sorted_workbook_bytes(), "application/octet-stream")}
    response = client.post(
        "/commands/group",
        data={"sheets": ["Getty Videos"], "sheets_present": "1"},
        files=files,
    )
    assert response.status_code == 200
    wb = load_workbook(io.BytesIO(response.content))
    header_row = [c.value for c in next(wb["Getty Videos"].iter_rows(min_row=1, max_row=1))]
    assert "Total Duration" in header_row
    ap_header_row = [c.value for c in next(wb["AP"].iter_rows(min_row=1, max_row=1))]
    assert "Total Duration" not in ap_header_row  # AP wasn't checked -> untouched


def test_group_every_checkbox_unchecked_with_marker_groups_nothing(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("ep1.xlsx", _sorted_workbook_bytes(), "application/octet-stream")}
    response = client.post(
        "/commands/group",
        data={"sheets_present": "1"},
        files=files,
    )
    assert response.status_code == 200
    wb = load_workbook(io.BytesIO(response.content))
    header_row = [c.value for c in next(wb["Getty Videos"].iter_rows(min_row=1, max_row=1))]
    assert "Total Duration" not in header_row


def test_group_no_marker_plain_form_keeps_default_sheets(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("ep1.xlsx", _sorted_workbook_bytes(), "application/octet-stream")}
    response = client.post("/commands/group", data={}, files=files)
    assert response.status_code == 200
    wb = load_workbook(io.BytesIO(response.content))
    header_row = [c.value for c in next(wb["Getty Videos"].iter_rows(min_row=1, max_row=1))]
    assert "Total Duration" in header_row  # DEFAULT_SHEETS still applies
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_web_routes_group.py -v`
Expected: FAIL — `parse_field`/`_build_kwargs` don't special-case `"sheet_checklist"` yet (checkboxes named `sheets` won't be collected as a list), and there's no `sheets_present` handling, so the "everything unchecked" case falls back to `DEFAULT_SHEETS` instead of grouping nothing.

- [ ] **Step 3: Implement**

In `web/files.py`, change `parse_field`'s type checks:

```python
def parse_field(field, raw):
    if field.type == "checkbox":
        return bool(raw)
    if raw is None or raw == "" or raw == []:
        return None if field.type in ("list", "sheet_checklist", "number") else (field.default or "")
    if field.type in ("list", "sheet_checklist"):
        if isinstance(raw, list):
            parts = [p.strip() for p in raw if p.strip()]
        else:
            parts = [p.strip() for p in raw.split(",") if p.strip()]
        return parts or None
    if field.type == "number":
        num = float(raw)
        return int(num) if num.is_integer() else num
    return raw
```

In `web/main.py`, change `_build_kwargs`:

```python
def _build_kwargs(form, fields) -> dict:
    """Read one value per field from a submitted form, using getlist() for
    "list"/"sheet_checklist" fields so multiple same-named inputs
    (checkboxes) all survive instead of only the first/last one."""
    return {
        f.name: parse_field(f, form.getlist(f.name) if f.type in ("list", "sheet_checklist") else form.get(f.name))
        for f in fields
    }
```

In `web/main.py`, in `command_submit`, right after the existing `sort`/`categories_present` block, add the `group` equivalent:

```python
    if (
        spec.slug == "sort"
        and form.get("categories_present")
        and not kwargs.get("categories")
        and not kwargs.get("skip_categories")
    ):
        kwargs["skip_categories"] = list(SORT_CATEGORIES)

    # Mirrors the categories_present marker above: group's sheet checklist
    # auto-unchecks nothing on its own, but a user who deliberately
    # unchecks every box should get "group nothing," not silently fall
    # back to DEFAULT_SHEETS (parse_field turns an empty submitted list
    # into None, indistinguishable from "field omitted" without this).
    if spec.slug == "group" and form.get("sheets_present") and not kwargs.get("sheets"):
        kwargs["sheets"] = []
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_web_routes_group.py tests/test_web_commands.py -v`
Expected: PASS.

- [ ] **Step 5: Run the full suite to confirm nothing else regressed**

Run: `pytest -q`
Expected: PASS, all tests green (sort's `categories_present` tests in particular, since that block was edited).

- [ ] **Step 6: Commit**

```bash
git add web/files.py web/main.py tests/test_web_routes_group.py
git commit -m "feat: sheet_checklist multi-value parsing + group's sheets_present marker"
```

---

## Task 7: Template — `sheet_checklist` fieldset + optional-sheet dropdown marker

**Files:**
- Modify: `web/templates/command.html`
- Test: `tests/test_web_commands.py`

**Interfaces:**
- Consumes: `FieldSpec.type == "sheet_checklist"`, `FieldSpec.allow_missing_sheet` (Task 4).
- Produces: rendered HTML — a `<fieldset data-analyze data-source="sheets" data-checklist-name="sheets">` (empty at render time, JS populates it in Task 8) for `sheet_checklist` fields; a `data-optional="true"` attribute on the `<select>` for any `options_source="sheets"` field with `allow_missing_sheet=True`.

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_web_commands.py` (uses a logged-in `TestClient` the way `test_web_routes_*` files do — check the top of `tests/test_web_commands.py` for its existing login helper and reuse it; if the file doesn't already have one, add the same `_logged_in_client(monkeypatch)` helper used in `tests/test_web_routes_getty_ids.py`):

```python
def test_group_page_renders_sheet_checklist_fieldset(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/group")
    assert response.status_code == 200
    body = response.text
    assert 'data-checklist-name="sheets"' in body
    assert 'name="sheets_present"' in body


def test_getty_ids_page_renders_sheet_name_dropdowns(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/getty-ids")
    assert response.status_code == 200
    body = response.text
    assert 'name="sheet_name"' in body and 'data-source="sheets"' in body
    assert 'data-optional="true"' in body  # stills_sheet_name only


def test_fu_grid_page_still_renders_plain_sheets_select(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/fu-grid")
    assert response.status_code == 200
    assert 'name="sheet"' in response.text
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pytest tests/test_web_commands.py -k "checklist_fieldset or sheet_name_dropdowns or plain_sheets_select" -v`
Expected: FAIL — no `sheet_checklist` branch in the template yet, no `data-optional` attribute.

- [ ] **Step 3: Implement**

In `web/templates/command.html`, the field-rendering loop currently reads:

```html
    {% for f in spec.fields %}
    {% if f.name == "categories" and spec.slug == "sort" %}
    <fieldset class="categories">
      ...
    </fieldset>
    {% elif f.options_source %}
    <label>
      {{ f.label }}{% if f.required %} (required){% endif %}
      {% set current = (values or {}).get(f.name, f.default) %}
      <select name="{{ f.name }}" data-analyze data-source="{{ f.options_source }}">
        <option value="{{ current if current is not none else '' }}" selected>{{ current if current is not none else '(pick a file)' }}</option>
      </select>
    </label>
    {% else %}
    ...
    {% endif %}
    {% endfor %}
```

Add a new `elif` branch for `sheet_checklist` between the `categories` branch and the generic `options_source` branch, and add the `data-optional` attribute to the generic `<select>` branch:

```html
    {% for f in spec.fields %}
    {% if f.name == "categories" and spec.slug == "sort" %}
    <fieldset class="categories">
      ...
    </fieldset>
    {% elif f.type == "sheet_checklist" %}
    <fieldset class="sheet-checklist" data-analyze data-source="sheets" data-checklist-name="{{ f.name }}">
      <legend>{{ f.label }}</legend>
      <input type="hidden" name="{{ f.name }}_present" value="1">
      <p class="hint">Upload a file to see its sheets.</p>
    </fieldset>
    {% elif f.options_source %}
    <label>
      {{ f.label }}{% if f.required %} (required){% endif %}
      {% set current = (values or {}).get(f.name, f.default) %}
      <select name="{{ f.name }}" data-analyze data-source="{{ f.options_source }}"
              {% if f.allow_missing_sheet %}data-optional="true"{% endif %}>
        <option value="{{ current if current is not none else '' }}" selected>{{ current if current is not none else '(pick a file)' }}</option>
      </select>
    </label>
    {% else %}
    ...
    {% endif %}
    {% endfor %}
```

(The `{% else %}` branch and everything outside this loop is untouched.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `pytest tests/test_web_commands.py -v`
Expected: PASS — new tests, plus every pre-existing template-rendering test in the file.

- [ ] **Step 5: Commit**

```bash
git add web/templates/command.html tests/test_web_commands.py
git commit -m "feat: render sheet_checklist fieldset and data-optional marker in command.html"
```

---

## Task 8: `analyze.js` — populate the checklist, honor the optional-select marker

No JS test harness exists in this repo (see Global Constraints) — this task is verified through the route tests already added (Tasks 5/6 assert on the JSON `suggestions`/`warnings` the JS consumes, Task 7 asserts the markup the JS attaches to) plus a manual browser check in Task 9.

**Files:**
- Modify: `web/static/analyze.js`

**Interfaces:**
- Consumes: `data.sheets` (existing `{name, rows}` list from the analyze JSON response), `data.suggestions[fieldName]` (a list for a `sheet_checklist` field, a string for a plain `sheets`/`headers` select — unchanged contract, per Task 5).
- Produces: `rebuildChecklist(fieldset, options, suggested, name)`, wired into `apply()`; `rebuildSelect` gains an optional 5th `labels` parameter for the "(none)" display text.

- [ ] **Step 1: Add `rebuildChecklist` and extend `rebuildSelect`**

In `web/static/analyze.js`, change the `rebuildSelect` signature and body to accept an optional label map:

```javascript
  // Rebuild a <select>'s options from `options`, keeping a valid current
  // selection or honoring a server suggestion. When neither is possible the
  // value is silently substituted to options[0]; push a note into `subs` so
  // apply() can surface it in the warning banner (spec decision 6: warn,
  // fall back, keep Run enabled). `labels`, when given, maps an option
  // value to its display text (used for the "(none)" blank option on an
  // optional sheet field -- the value stays "" but shouldn't read blank).
  function rebuildSelect(sel, options, suggested, subs, labels) {
    var prev = sel.value;
    var want;
    if (options.indexOf(suggested) !== -1) {
      want = suggested;
    } else if (options.indexOf(prev) !== -1) {
      want = prev;
    } else {
      want = options[0];
      if (prev && prev !== want && subs) {
        subs.push(fieldLabel(sel) + " '" + prev + "' not in this file — using '" + want + "'");
      }
    }
    sel.innerHTML = "";
    options.forEach(function (opt) {
      var o = document.createElement("option");
      o.value = opt;
      o.textContent = (labels && labels[opt] !== undefined) ? labels[opt] : opt;
      if (opt === want) o.selected = true;
      sel.appendChild(o);
    });
  }

  // Rebuild a dynamic checkbox list (group's "sheets to group" field) from
  // `options` (real sheet names). Preserves any currently-checked box whose
  // value is still in `options`; otherwise checks exactly the boxes named
  // in `suggested`. There is no analog of rebuildSelect's "fall back to
  // options[0]" here -- an empty checklist (nothing suggested, nothing
  // previously checked) is a valid state, unlike a required single-select.
  function rebuildChecklist(fieldset, options, suggested, name) {
    var hadAny = false;
    var prevChecked = {};
    fieldset.querySelectorAll('input[type=checkbox]').forEach(function (box) {
      hadAny = true;
      if (box.checked) prevChecked[box.value] = true;
    });
    var wantChecked = hadAny ? prevChecked : (suggested || []).reduce(function (acc, opt) {
      acc[opt] = true;
      return acc;
    }, {});

    fieldset.querySelectorAll("label.sheet-checkbox").forEach(function (el) { el.remove(); });
    options.forEach(function (opt) {
      var label = document.createElement("label");
      label.className = "sheet-checkbox";
      var box = document.createElement("input");
      box.type = "checkbox";
      box.name = name;
      box.value = opt;
      box.checked = !!wantChecked[opt];
      label.appendChild(box);
      label.appendChild(document.createTextNode(" " + opt));
      fieldset.appendChild(label);
    });
  }
```

- [ ] **Step 2: Wire both into `apply()`**

Change the `apply()` function's `sheets`-select block from:

```javascript
    if (sheetNames.length) {
      form.querySelectorAll('select[data-source="sheets"]').forEach(function (sel) {
        rebuildSelect(sel, sheetNames, suggestions[sel.name], subs);
      });
    }
```

to:

```javascript
    if (sheetNames.length) {
      form.querySelectorAll('select[data-source="sheets"]').forEach(function (sel) {
        var optional = sel.dataset.optional === "true";
        var opts = optional ? [""].concat(sheetNames) : sheetNames;
        var labels = optional ? {"": "(none)"} : null;
        rebuildSelect(sel, opts, suggestions[sel.name], subs, labels);
      });
      form.querySelectorAll("fieldset[data-checklist-name]").forEach(function (fieldset) {
        var name = fieldset.getAttribute("data-checklist-name");
        rebuildChecklist(fieldset, sheetNames, suggestions[name], name);
      });
    }
```

- [ ] **Step 3: Trigger re-analysis when a checklist box is toggled**

At the bottom of the IIFE, where existing `select[data-analyze]` elements get a `change` listener:

```javascript
  fileInput.addEventListener("change", function () { analyze(false); });
  form.querySelectorAll("select[data-analyze]").forEach(function (sel) {
    sel.addEventListener("change", function () { analyze(false); });
  });
```

add, right after it:

```javascript
  // Checkboxes inside a checklist fieldset are created/destroyed by
  // rebuildChecklist on every analyze response, so a listener bound to one
  // instance wouldn't survive the next rebuild -- delegate on the fieldset
  // itself instead (matches the "no per-checkbox listeners" note in the
  // design doc).
  form.querySelectorAll("fieldset[data-checklist-name]").forEach(function (fieldset) {
    fieldset.addEventListener("change", function (e) {
      if (e.target && e.target.type === "checkbox") analyze(false);
    });
  });
```

- [ ] **Step 4: Add checklist styling**

In `web/static/style.css`, alongside the existing `.category-checkbox`/`.categories` rules, add:

```css
.sheet-checklist { display: flex; flex-wrap: wrap; gap: 0.5rem 1.5rem; border: 1px solid #ccc; padding: 0.5rem 0.75rem; }
.sheet-checkbox { display: flex; align-items: center; gap: 0.4rem; width: 220px; margin-top: 0.25rem; }
.sheet-checklist .hint { flex-basis: 100%; color: #666; font-size: 0.9em; margin: 0.5rem 0 0; }
```

- [ ] **Step 5: Run the full Python test suite (JS has no runnable test harness, but nothing here should touch Python behavior)**

Run: `pytest -q`
Expected: PASS, unchanged count from before this task (JS/CSS files aren't imported by any Python test).

- [ ] **Step 6: Commit**

```bash
git add web/static/analyze.js web/static/style.css
git commit -m "feat: analyze.js populates the sheet checklist and honors optional sheet selects"
```

---

## Task 9: Full regression run + manual smoke check

**Files:** none modified — verification only.

- [ ] **Step 1: Full test suite**

Run: `pytest -q`
Expected: PASS, all tests green (444 at plan start + every test added in Tasks 1–7).

- [ ] **Step 2: Manual smoke check — getty-ids sheet dropdowns**

Run: `uvicorn web.main:app --reload` (or however this project normally runs the dev server — check `README.md`'s local-dev section if unsure), log in, open `/commands/getty-ids`, upload two `.xlsx` files whose active tab is something other than "Getty Videos" (e.g. re-use the `data/in/14.09 - WR/in/WR EP1.xlsx`/`WR EP2.xlsx` files from this session's earlier work — their active tab is "COST"). Confirm:
- "Video sheet name" and "Stills sheet name" render as dropdowns (not text inputs), auto-selecting "Getty Videos"/"(none)" (this project has no stills sheet).
- "Filename column" and "Seconds column" dropdowns show `Track`/`Clip Name`/`Seconds` — not `EP1`/`EP2`.
- No "columns differ" warning banner.

- [ ] **Step 3: Manual smoke check — group checklist**

Open `/commands/group`, upload the same files. Confirm the "Sheets to group" field renders as a checkbox list of the file's real sheets (AP, Getty Videos, COST, worksheet, artlist, shutterstock, CAM+GFX or similar), with the known category sheets pre-checked and the rest (COST, worksheet, ...) left unchecked. Toggle a checkbox and confirm the Filename/Duration column dropdowns update.

- [ ] **Step 4: Manual smoke check — fu-grid**

Open `/commands/fu-grid`, upload a sorted file. Confirm "Source sheet" still behaves as before (dropdown, defaults to "3rd parties") and the Filename/Duration column dropdowns now show that sheet's real headers.

- [ ] **Step 5: Confirm `sort`/`dedupe` are untouched**

Open `/commands/sort` and `/commands/dedupe`, upload a raw (unsorted) file, confirm both pages behave exactly as before this plan (no new dropdowns/checklists, same suggestions as previously observed).

No commit for this task — it's a verification pass. If any manual check fails, return to the relevant task above, fix, and re-run its tests before continuing.
