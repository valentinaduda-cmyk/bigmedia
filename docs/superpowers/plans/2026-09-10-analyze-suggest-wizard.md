# Analyze-and-Suggest Wizard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every web command follows one flow — upload a file, the app inspects it and pre-fills the form with suggestions, the user eyeballs and edits, then runs — built as a reusable framework and wired to Sort first.

**Architecture:** A generic `analyze_files()` reads the header union, sheet
names and row counts from the uploaded workbook(s). Each `CommandSpec` may
carry an optional `analyze` hook for command-specific extras (Sort's
per-category clip counts). One route `POST /commands/{slug}/analyze`
returns a JSON bundle of `headers`, `sheets`, `suggestions`, `annotations`
and `warnings`. The browser keeps the file in the `<input>` and re-sends it
on Run — the server stays stateless. A single generic script populates
`<select>` dropdowns and category checkboxes from the analyze response.

**Tech Stack:** Python 3.9+, FastAPI, Jinja2, openpyxl, pytest, vanilla JS
(no build step), `fastapi.testclient.TestClient`.

**Spec:** `docs/superpowers/specs/2026-09-10-analyze-suggest-wizard-design.md`

## Global Constraints

- Business logic modules under `bigmedia/` stay import-light; `analyze_files`
  goes in `bigmedia/xlsx_utils.py` alongside the other openpyxl helpers.
- New reader functions open workbooks `read_only=True, data_only=True` and
  call `wb.close()`, matching `count_categories`.
- No new runtime dependencies. `requirements.txt` and `pyproject.toml`
  unchanged.
- Web routes that touch data require `dependencies=[Depends(require_login)]`.
- Tests use the synthetic fixtures in `tests/conftest.py` or build
  workbooks inline with `openpyxl`; never commit real client data.
- Commit messages follow the repo style: `feat(web): ...`, `feat: ...`,
  `test: ...`, `docs: ...`. End every commit body with
  `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>`.
- Full `pytest` must stay green (394 passing before this plan).

---

### Task 1: `analyze_files()` — generic workbook inspection

**Files:**
- Modify: `bigmedia/xlsx_utils.py` (add function + module imports if needed)
- Test: `tests/test_xlsx_utils.py` (create)

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces:
  `analyze_files(paths: list[str]) -> dict` returning exactly these keys:
  ```python
  {
      "headers": list[str],   # union of row-1 values of each file's ACTIVE
                              # sheet, first-seen order, blank cells skipped
      "sheets": list[dict],   # [{"name": str, "rows": int}, ...] union over
                              # all files' sheet names; "rows" summed across
                              # files that have that sheet; header row excluded
      "warnings": list[str],
  }
  ```
  Warning strings:
  - `"<filename>: could not read (<exc>)"` when `load_workbook` raises;
    that file is skipped, the rest still processed.
  - `"<filename>: columns differ from <first-filename>"` when a later
    file's active-sheet header set differs from the first readable file's.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_xlsx_utils.py`:

```python
import pytest
from openpyxl import Workbook

from bigmedia.xlsx_utils import analyze_files


def _make(path, sheets):
    """sheets: dict of {title: [header_row, *data_rows]}."""
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title=title)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return str(path)


def test_headers_are_union_first_seen_order(tmp_path):
    a = _make(tmp_path / "a.xlsx", {"S": [["Clip Name", "Duration"], ["x.mov", 1]]})
    b = _make(tmp_path / "b.xlsx", {"S": [["Clip Name", "Notes"], ["y.mov", "n"]]})
    result = analyze_files([a, b])
    assert result["headers"] == ["Clip Name", "Duration", "Notes"]


def test_blank_header_cells_are_skipped(tmp_path):
    a = _make(tmp_path / "a.xlsx", {"S": [["Clip Name", None, "Duration"], ["x", 1, 2]]})
    assert analyze_files([a])["headers"] == ["Clip Name", "Duration"]


def test_sheets_list_names_and_row_counts_excluding_header(tmp_path):
    a = _make(tmp_path / "a.xlsx", {
        "Getty Videos": [["Clip Name"], ["x.mov"], ["y.mov"]],
        "Getty Stills": [["Clip Name"]],
    })
    sheets = {s["name"]: s["rows"] for s in analyze_files([a])["sheets"]}
    assert sheets == {"Getty Videos": 2, "Getty Stills": 0}


def test_sheet_row_counts_sum_across_files(tmp_path):
    a = _make(tmp_path / "a.xlsx", {"Data": [["Clip Name"], ["x"], ["y"]]})
    b = _make(tmp_path / "b.xlsx", {"Data": [["Clip Name"], ["z"]]})
    sheets = {s["name"]: s["rows"] for s in analyze_files([a, b])["sheets"]}
    assert sheets == {"Data": 3}


def test_disagreeing_columns_warn(tmp_path):
    a = _make(tmp_path / "master_a.xlsx", {"S": [["Clip Name", "Duration"], ["x", 1]]})
    b = _make(tmp_path / "master_b.xlsx", {"S": [["Name"], ["y"]]})
    warnings = analyze_files([a, b])["warnings"]
    assert any("master_b.xlsx" in w and "differ" in w for w in warnings)


def test_unreadable_file_warns_and_others_still_processed(tmp_path):
    good = _make(tmp_path / "good.xlsx", {"S": [["Clip Name"], ["x"]]})
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a real xlsx")
    result = analyze_files([str(bad), good])
    assert result["headers"] == ["Clip Name"]
    assert any("bad.xlsx" in w and "could not read" in w for w in result["warnings"])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_xlsx_utils.py -v`
Expected: FAIL — `ImportError: cannot import name 'analyze_files'`

- [ ] **Step 3: Implement `analyze_files`**

Add to `bigmedia/xlsx_utils.py` (top already has `import copy`; add
`from pathlib import Path` and `from openpyxl import load_workbook` if not
present — check the existing imports first and only add what's missing):

```python
def analyze_files(paths):
    """Read-only inspection of one or more clip-list workbooks, for the web
    UI's analyze-and-suggest step. Returns the union of the active sheets'
    header names, every sheet's name and (header-excluded) row count summed
    across files, and warnings for files that won't open or whose columns
    disagree with the first readable file. Never raises for a bad file --
    it lands in "warnings" and the others are still processed."""
    headers = []
    seen_headers = set()
    first_header_set = None
    first_name = None
    sheet_rows = {}
    sheet_order = []
    warnings = []

    for path in paths:
        name = Path(path).name
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:  # openpyxl raises several unrelated types
            warnings.append(f"{name}: could not read ({exc})")
            continue

        active = wb.active
        file_headers = [
            c.value for c in next(active.iter_rows(min_row=1, max_row=1), [])
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
            rows = ws.max_row
            if rows is None:
                rows = sum(1 for _ in ws.iter_rows())
            count = max(rows - 1, 0)
            if title not in sheet_rows:
                sheet_rows[title] = 0
                sheet_order.append(title)
            sheet_rows[title] += count

        wb.close()

    return {
        "headers": headers,
        "sheets": [{"name": t, "rows": sheet_rows[t]} for t in sheet_order],
        "warnings": warnings,
    }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_xlsx_utils.py -v`
Expected: PASS (6 tests)

- [ ] **Step 5: Run the full suite**

Run: `python -m pytest -q`
Expected: PASS (400 tests)

- [ ] **Step 6: Commit**

```bash
git add bigmedia/xlsx_utils.py tests/test_xlsx_utils.py
git commit -m "$(cat <<'EOF'
feat: add analyze_files() for read-only workbook inspection

Header union, per-sheet row counts, and warnings for unreadable or
column-mismatched files. Feeds the web UI's upcoming analyze step.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 2: `analyze_sort()` — Sort's per-command analyzer

**Files:**
- Modify: `bigmedia/sort_workbook.py` (add function below `count_categories`)
- Test: `tests/test_sort_workbook.py` (append tests)

**Interfaces:**
- Consumes: existing `count_categories(paths, name_column="Clip Name") -> dict`.
- Produces:
  `analyze_sort(paths: list[str], name_column: str = "Clip Name") -> dict`
  returning:
  - on success:
    `{"suggestions": {"categories": [<cats with count > 0>]},
      "annotations": {"categories": {<cat>: <count>, ...}}}`
    where `annotations.categories` has every entry of `CATEGORY_ORDER`.
  - when the name column is absent from a file (`count_categories` raises
    `ValueError`): `{"warnings": [<str(exc)>]}` and nothing else.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_sort_workbook.py` (it already imports from
`bigmedia.sort_workbook` and uses `openpyxl`; match the file's existing
import style):

```python
def test_analyze_sort_suggests_only_nonzero_categories(tmp_path):
    from bigmedia.sort_workbook import analyze_sort
    wb = Workbook()
    ws = wb.active
    ws.append(["Clip Name"])
    ws.append(["BM1234_x.mxf"])       # AP
    ws.append(["shutterstock_9.mp4"]) # Shutterstock
    path = tmp_path / "m.xlsx"
    wb.save(path)

    result = analyze_sort([str(path)])
    assert set(result["suggestions"]["categories"]) == {"AP", "Shutterstock"}
    assert result["annotations"]["categories"]["AP"] == 1
    assert result["annotations"]["categories"]["Reuters"] == 0
    assert "warnings" not in result


def test_analyze_sort_missing_name_column_warns_not_raises(tmp_path):
    from bigmedia.sort_workbook import analyze_sort
    wb = Workbook()
    wb.active.append(["Something Else"])
    path = tmp_path / "m.xlsx"
    wb.save(path)

    result = analyze_sort([str(path)], name_column="Clip Name")
    assert result["warnings"]
    assert "suggestions" not in result
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_sort_workbook.py -k analyze_sort -v`
Expected: FAIL — `ImportError: cannot import name 'analyze_sort'`

- [ ] **Step 3: Implement `analyze_sort`**

Add to `bigmedia/sort_workbook.py` directly after `count_categories`:

```python
def analyze_sort(paths, name_column="Clip Name"):
    """Analyze step for the web UI's Sort form: classify every clip without
    writing output and report which categories are worth keeping (a
    suggestion) plus the full per-category count (a display annotation). A
    name column that isn't in the file is a warning, not an error -- the
    user can pick the right column from the dropdown and re-run analysis."""
    try:
        counts = count_categories(paths, name_column=name_column)
    except ValueError as exc:
        return {"warnings": [str(exc)]}
    return {
        "suggestions": {"categories": [c for c, n in counts.items() if n > 0]},
        "annotations": {"categories": counts},
    }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_sort_workbook.py -k analyze_sort -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add bigmedia/sort_workbook.py tests/test_sort_workbook.py
git commit -m "$(cat <<'EOF'
feat: add analyze_sort() for the Sort form's analyze step

Wraps count_categories: suggests non-empty categories, annotates every
category with its clip count, downgrades a missing name column to a
warning.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 3: `FieldSpec.options_source` + `CommandSpec.analyze`

**Files:**
- Modify: `web/commands.py`
- Test: `tests/test_web_commands.py` (append tests)

**Interfaces:**
- Consumes: `analyze_sort` from Task 2.
- Produces:
  - `FieldSpec` gains `options_source: str | None = None` (values used:
    `"headers"`, `"sheets"`).
  - `CommandSpec` gains `analyze: Callable | None = None`.
  - `_NAME_COLUMN` and `_DURATION_COLUMN` module constants gain
    `options_source="headers"`.
  - `COMMANDS["fu-grid"]`'s `sheet` field gains `options_source="sheets"`.
    Group's `sheets` field is left alone — it is a multi-value `list` field
    and a single `<select>` would regress it; it gets `options_source` when
    group's own analyzer lands (out of scope here).
  - `COMMANDS["sort"].analyze is analyze_sort`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_web_commands.py`:

```python
from bigmedia.sort_workbook import analyze_sort


def test_name_and_duration_columns_are_header_dropdowns():
    for slug in ("sort", "group", "fu-grid"):
        by_name = {f.name: f for f in COMMANDS[slug].fields}
        assert by_name["name_column"].options_source == "headers"
    assert {f.name: f for f in COMMANDS["group"].fields}["duration_column"].options_source == "headers"
    # getty-ids and compare share _NAME_COLUMN too — they get the header dropdown for free
    assert {f.name: f for f in COMMANDS["getty-ids"].fields}["name_column"].options_source == "headers"


def test_sheet_fields_are_sheet_dropdowns():
    assert {f.name: f for f in COMMANDS["fu-grid"].fields}["sheet"].options_source == "sheets"
    # group's multi-value "sheets" field stays a text field until group's analyzer lands
    assert {f.name: f for f in COMMANDS["group"].fields}["sheets"].options_source is None


def test_only_sort_has_an_analyzer_for_now():
    assert COMMANDS["sort"].analyze is analyze_sort
    for slug in ("dedupe", "group", "fu-grid", "compare", "fix-getty", "getty-ids"):
        assert COMMANDS[slug].analyze is None


def test_fields_without_options_source_default_to_none():
    assert {f.name: f for f in COMMANDS["sort"].fields}["autofit"].options_source is None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_web_commands.py -k "options_source or analyzer or header_dropdown or sheet_dropdown" -v`
Expected: FAIL — `AttributeError: 'FieldSpec' object has no attribute 'options_source'`

- [ ] **Step 3: Implement the dataclass fields and wiring**

In `web/commands.py`:

1. Add the import near the other `bigmedia` imports:
   ```python
   from bigmedia.sort_workbook import sort_workbook, analyze_sort
   ```
   (replace the existing `from bigmedia.sort_workbook import sort_workbook` line)

2. Extend `FieldSpec`:
   ```python
   @dataclass
   class FieldSpec:
       name: str
       label: str
       type: str  # "text" | "number" | "checkbox" | "list"
       default: Any = None
       required: bool = False
       options_source: str = None  # None | "headers" | "sheets"
   ```

3. Extend `CommandSpec`:
   ```python
   @dataclass
   class CommandSpec:
       slug: str
       title: str
       upload_mode: str  # "batch" | "combine" | "pair" | "fix_getty"
       func: Callable
       fields: list
       output_suffix: str
       analyze: Callable = None
   ```

4. Update the shared column-field constants:
   ```python
   _NAME_COLUMN = FieldSpec("name_column", "Filename column", "text", "Clip Name", options_source="headers")
   _DURATION_COLUMN = FieldSpec("duration_column", "Duration column", "text", "Clip Duration", options_source="headers")
   ```

5. In `COMMANDS["sort"]`, add `analyze=analyze_sort,` to the `CommandSpec(...)` call.

6. Leave `COMMANDS["group"]`'s `sheets` field unchanged (multi-value list;
   deferred — see Interfaces note).

7. In `COMMANDS["fu-grid"]`, change the `sheet` field to:
   ```python
   FieldSpec("sheet", "Source sheet", "text", "3rd parties", options_source="sheets"),
   ```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_web_commands.py -v`
Expected: PASS (all, including the pre-existing `test_sort_fields_match_cli_options`)

- [ ] **Step 4b: Run the full suite — catch breakage in other web test files**

Run: `python -m pytest -q`
Expected: PASS. If `tests/test_web_final_fixes.py` or `tests/test_web_presets_ui.py`
assert on `FieldSpec` internals (e.g. `_NAME_COLUMN.type`), only the new
`options_source` attribute changed — update those assertions to keep their
intent and note it in the commit body. Template-level assertions there are
not affected yet (the template changes land in Task 6).

- [ ] **Step 5: Commit**

```bash
git add web/commands.py tests/test_web_commands.py
git commit -m "$(cat <<'EOF'
feat(web): add options_source to FieldSpec and analyze hook to CommandSpec

Marks column/sheet fields as analyze-populated dropdowns; lets a command
carry a per-command analyzer. Sort wired to analyze_sort.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 4: `run_analysis()` — merge generic + per-command analysis

**Files:**
- Create: `web/analyze.py`
- Test: `tests/test_web_analyze.py` (create)

**Interfaces:**
- Consumes: `analyze_files` (Task 1), `CommandSpec.analyze` (Task 3).
- Produces:
  `run_analysis(spec, paths: list[str], form) -> dict` where `form` is a
  mapping (a Starlette `FormData` or a plain `dict`) supporting `.get(key)`.
  Returns exactly:
  ```python
  {
      "headers": list[str],
      "sheets": list[dict],
      "suggestions": dict,
      "annotations": dict,
      "warnings": list[str],
  }
  ```
  Generic `headers`/`sheets`/`warnings` come from `analyze_files`. If
  `spec.analyze` is set it is called as
  `spec.analyze(paths, **kwargs)` where `kwargs` is every non-`paths`
  parameter of its signature that has a truthy value in `form`; its
  `suggestions`/`annotations` are merged in and its `warnings` appended.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_web_analyze.py`:

```python
from openpyxl import Workbook

from web.analyze import run_analysis
from web.commands import COMMANDS


def _sheet(path, headers, names):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for n in names:
        ws.append([n])
    wb.save(path)
    return str(path)


def test_generic_only_command_returns_headers_and_sheets(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Clip Name"], ["x.mov"])
    result = run_analysis(COMMANDS["dedupe"], [path], {})
    assert result["headers"] == ["Clip Name"]
    assert result["sheets"][0]["name"]
    assert result["suggestions"] == {}
    assert result["annotations"] == {}


def test_sort_merges_analyzer_suggestions_and_annotations(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Clip Name"], ["BM1234_x.mxf", "random_xyz.mov"])
    result = run_analysis(COMMANDS["sort"], [path], {"name_column": "Clip Name"})
    assert "AP" in result["suggestions"]["categories"]
    assert result["annotations"]["categories"]["AP"] == 1


def test_sort_analyzer_uses_form_name_column(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Media File"], ["BM1234_x.mxf"])
    result = run_analysis(COMMANDS["sort"], [path], {"name_column": "Media File"})
    assert result["annotations"]["categories"]["AP"] == 1
    assert not result["warnings"]


def test_sort_analyzer_warning_is_surfaced(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Media File"], ["BM1234_x.mxf"])
    result = run_analysis(COMMANDS["sort"], [path], {"name_column": "Clip Name"})
    assert result["warnings"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_web_analyze.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'web.analyze'`

- [ ] **Step 3: Implement `web/analyze.py`**

```python
import inspect

from bigmedia.xlsx_utils import analyze_files


def _analyzer_kwargs(func, form):
    """Every parameter the analyzer declares (besides `paths`) that has a
    truthy value in the submitted form, so the analyzer sees the user's
    current dropdown choices (e.g. the name column) rather than only its
    own defaults."""
    kwargs = {}
    for pname, param in inspect.signature(func).parameters.items():
        if pname == "paths" or param.kind in (param.VAR_POSITIONAL, param.VAR_KEYWORD):
            continue
        value = form.get(pname)
        if value:
            kwargs[pname] = value
    return kwargs


def run_analysis(spec, paths, form):
    generic = analyze_files(paths)
    result = {
        "headers": generic["headers"],
        "sheets": generic["sheets"],
        "suggestions": {},
        "annotations": {},
        "warnings": list(generic["warnings"]),
    }
    if spec.analyze:
        extra = spec.analyze(paths, **_analyzer_kwargs(spec.analyze, form))
        result["suggestions"].update(extra.get("suggestions", {}))
        result["annotations"].update(extra.get("annotations", {}))
        result["warnings"].extend(extra.get("warnings", []))
    return result
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_web_analyze.py -v`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add web/analyze.py tests/test_web_analyze.py
git commit -m "$(cat <<'EOF'
feat(web): add run_analysis() merging generic and per-command analysis

Combines analyze_files() output with an optional CommandSpec.analyze
hook, passing the user's current form values into the analyzer.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 5: Generic `POST /commands/{slug}/analyze` route

**Files:**
- Modify: `web/main.py` (remove `sort_analyze`, add `command_analyze`;
  update the import from `web.commands`)
- Test: `tests/test_web_routes_sort_analyze.py` (rework)

**Interfaces:**
- Consumes: `run_analysis` (Task 4).
- Produces: `POST /commands/{slug}/analyze`, login-gated, body =
  `files` (one or more `UploadFile`) plus any form fields. Returns:
  - `400 {"error": "Not an .xlsx file: ..."}` for non-xlsx uploads
    (unchanged behaviour).
  - `400 {"error": "<message>"}` if `run_analysis` raises.
  - `200` with the `run_analysis` dict otherwise.
  - `303` redirect when not logged in (unchanged).

- [ ] **Step 1: Rework the test file**

Replace `tests/test_web_routes_sort_analyze.py` with:

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


def _xlsx_bytes(names, header="Clip Name", extra_sheets=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append([header])
    for name in names:
        ws.append([name])
    for title, rows in (extra_sheets or {}).items():
        s = wb.create_sheet(title=title)
        for row in rows:
            s.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_analyze_returns_category_annotations(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _xlsx_bytes(["BM1234_x.mxf", "BM5678_y.mxf", "random_file_xyz.mov"])
    files = {"files": ("master.xlsx", content, "application/octet-stream")}
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    body = response.json()
    assert body["annotations"]["categories"]["AP"] == 2
    assert body["annotations"]["categories"]["3rd parties"] == 1
    assert "AP" in body["suggestions"]["categories"]
    assert "Reuters" not in body["suggestions"]["categories"]


def test_analyze_returns_headers_and_sheets(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _xlsx_bytes(["x.mov"], extra_sheets={"Getty Videos": [["Clip Name"], ["g.mov"]]})
    files = {"files": ("master.xlsx", content, "application/octet-stream")}
    response = client.post("/commands/sort/analyze", data={}, files=files)
    body = response.json()
    assert body["headers"] == ["Clip Name"]
    names = {s["name"]: s["rows"] for s in body["sheets"]}
    assert names == {"Master": 1, "Getty Videos": 1}


def test_analyze_union_across_multiple_files(monkeypatch):
    client = _logged_in_client(monkeypatch)
    a = _xlsx_bytes(["BM1_x.mxf"], header="Clip Name")
    b = _xlsx_bytes(["BM2_y.mxf"], header="Name")
    files = [
        ("files", ("a.xlsx", a, "application/octet-stream")),
        ("files", ("b.xlsx", b, "application/octet-stream")),
    ]
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    body = response.json()
    assert set(body["headers"]) == {"Clip Name", "Name"}
    assert body["annotations"]["categories"]["AP"] == 2
    assert any("differ" in w for w in body["warnings"])


def test_analyze_missing_name_column_warns_but_200(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _xlsx_bytes(["BM1234_x.mxf"], header="Media File")
    files = {"files": ("master.xlsx", content, "application/octet-stream")}
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    assert response.json()["warnings"]


def test_analyze_generic_command_has_no_suggestions(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _xlsx_bytes(["x.mov"])
    files = {"files": ("master.xlsx", content, "application/octet-stream")}
    response = client.post("/commands/dedupe/analyze", data={}, files=files)
    body = response.json()
    assert body["headers"] == ["Clip Name"]
    assert body["suggestions"] == {}


def test_analyze_unknown_slug_404(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("m.xlsx", _xlsx_bytes(["x"]), "application/octet-stream")}
    response = client.post("/commands/nope/analyze", data={}, files=files)
    assert response.status_code == 404


def test_analyze_rejects_non_xlsx(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("notes.txt", b"hello", "text/plain")}
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 400
    assert "notes.txt" in response.json()["error"]


def test_analyze_requires_login(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    client = TestClient(app, follow_redirects=False)
    files = {"files": ("master.xlsx", _xlsx_bytes(["x.mov"]), "application/octet-stream")}
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 303


def test_sort_run_accepts_multiple_categories_checkboxes(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _xlsx_bytes([
        "BM1234_something.mxf",                  # AP
        "045_SO_EP18_01_3DExplainer_TXLS.mov",   # GFX
        "shutterstock_777.mp4",                  # Shutterstock
    ])
    files = {"files": ("master.xlsx", content, "application/octet-stream")}
    response = client.post(
        "/commands/sort",
        data={"name_column": "Clip Name", "categories": ["AP", "GFX"]},
        files=files,
    )
    assert response.status_code == 200
    wb = load_workbook(io.BytesIO(response.content))
    assert "AP" in wb.sheetnames
    assert "GFX" in wb.sheetnames
    assert "Shutterstock" not in wb.sheetnames
    assert "3rd parties" in wb.sheetnames
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_web_routes_sort_analyze.py -v`
Expected: FAIL — the new `annotations`/`headers` assertions fail against
the old `sort_analyze` route that returns `{"counts": ...}`; the
`/commands/dedupe/analyze` test 404s or errors.

- [ ] **Step 3: Swap the route in `web/main.py`**

1. Update the `web.commands` import line to add `run_analysis`? No —
   `run_analysis` lives in `web.analyze`. Add a new import near the other
   `web.` imports:
   ```python
   from web.analyze import run_analysis
   ```

2. Delete the entire `sort_analyze` function (the
   `@app.post("/commands/sort/analyze", ...)` block, roughly lines
   158-175) and the now-unused `count_categories` import
   (`from bigmedia.sort_workbook import count_categories`) if nothing else
   in the file uses it — grep first; `group_duplicates_workbook` import
   stays.

3. Add, directly above `command_submit`:
   ```python
   @app.post("/commands/{slug}/analyze", dependencies=[Depends(require_login)])
   async def command_analyze(request: Request, slug: str, files: List[UploadFile] = File(...)):
       spec = _command_or_404(slug)
       bad = reject_non_xlsx([f.filename for f in files])
       if bad:
           return JSONResponse({"error": f"Not an .xlsx file: {', '.join(bad)}"}, status_code=400)
       form = await request.form()
       with tempfile.TemporaryDirectory() as tmp:
           paths = _save_uploads(files, Path(tmp))
           try:
               data = run_analysis(spec, [str(p) for p in paths], form)
           except Exception as exc:
               logger.exception("analyze failed for %s", slug)
               return JSONResponse({"error": str(exc)}, status_code=400)
       return JSONResponse(data)
   ```

- [ ] **Step 4: Run the reworked route tests**

Run: `python -m pytest tests/test_web_routes_sort_analyze.py -v`
Expected: PASS (10 tests)

- [ ] **Step 5: Run the full suite**

Run: `python -m pytest -q`
Expected: PASS (~415 tests, no failures)

- [ ] **Step 6: Commit**

```bash
git add web/main.py tests/test_web_routes_sort_analyze.py
git commit -m "$(cat <<'EOF'
feat(web): generic /commands/{slug}/analyze route

Replaces the Sort-only analyze endpoint. Any command can be analyzed;
the response carries headers, sheets, suggestions, annotations and
warnings via run_analysis().

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 6: Template dropdowns, generic analyze script, warning banner

**Files:**
- Create: `web/static/analyze.js`
- Modify: `web/templates/command.html`
- Modify: `web/static/style.css`
- Test: `tests/test_web_command_form.py` (create)

**Interfaces:**
- Consumes: `FieldSpec.options_source` (Task 3), the
  `/commands/{slug}/analyze` route (Task 5).
- Produces: rendered `command.html` where
  - `<form>` carries `data-analyze-url="/commands/{slug}/analyze"`;
  - each `options_source` field renders as
    `<select name="{name}" data-analyze data-source="{headers|sheets}">`
    with a single `<option>` for the current/default value;
  - a `<div id="analyze-warnings" class="warn-banner" hidden>` sits at the
    top of the options `<fieldset>`;
  - `<script src="/static/analyze.js"></script>` is included once, for
    every command (not Sort-gated).

- [ ] **Step 1: Write the failing tests**

Create `tests/test_web_command_form.py`:

```python
from fastapi.testclient import TestClient

from web.auth import hash_password
from web.main import app


def _client(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    c = TestClient(app)
    c.post("/login", data={"password": "pw"})
    return c


def test_name_column_renders_as_analyze_select(monkeypatch):
    html = _client(monkeypatch).get("/commands/dedupe").text
    assert '<select name="name_column"' in html
    assert 'data-analyze' in html
    assert 'data-source="headers"' in html


def test_form_has_analyze_url_and_script(monkeypatch):
    html = _client(monkeypatch).get("/commands/dedupe").text
    assert 'data-analyze-url="/commands/dedupe/analyze"' in html
    assert '/static/analyze.js' in html


def test_warning_banner_present_and_hidden(monkeypatch):
    html = _client(monkeypatch).get("/commands/sort").text
    assert 'id="analyze-warnings"' in html
    assert 'hidden' in html.split('id="analyze-warnings"')[1][:40]


def test_sort_still_has_category_checkboxes(monkeypatch):
    html = _client(monkeypatch).get("/commands/sort").text
    assert 'name="categories"' in html
    assert 'class="category-name"' in html


def test_fu_grid_sheet_field_is_sheet_select(monkeypatch):
    html = _client(monkeypatch).get("/commands/fu-grid").text
    assert '<select name="sheet"' in html
    assert 'data-source="sheets"' in html


def test_group_sheets_field_stays_text_for_now(monkeypatch):
    # group's multi-value "sheets" field is deferred until group's analyzer
    html = _client(monkeypatch).get("/commands/group").text
    assert '<input type="text" name="sheets"' in html
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_web_command_form.py -v`
Expected: FAIL — current template renders `<input type="text" name="name_column">`, no select, no `analyze.js`.

- [ ] **Step 3: Create `web/static/analyze.js`**

```javascript
(function () {
  var form = document.querySelector("form[data-analyze-url]");
  if (!form) return;
  var url = form.getAttribute("data-analyze-url");
  var fileInput = form.querySelector('input[type="file"]');
  if (!fileInput) return;
  var banner = document.getElementById("analyze-warnings");

  function payload() {
    var fd = new FormData();
    for (var i = 0; i < fileInput.files.length; i++) fd.append("files", fileInput.files[i]);
    form.querySelectorAll("input[name], select[name]").forEach(function (el) {
      if (el.type === "file") return;
      if (el.type === "checkbox") { if (el.checked) fd.append(el.name, el.value || "on"); }
      else fd.append(el.name, el.value);
    });
    return fd;
  }

  function rebuildSelect(sel, options, suggested) {
    var want = options.indexOf(suggested) !== -1 ? suggested
             : (options.indexOf(sel.value) !== -1 ? sel.value : options[0]);
    sel.innerHTML = "";
    options.forEach(function (opt) {
      var o = document.createElement("option");
      o.value = opt;
      o.textContent = opt;
      if (opt === want) o.selected = true;
      sel.appendChild(o);
    });
  }

  function apply(data) {
    var suggestions = data.suggestions || {};
    var sheetNames = (data.sheets || []).map(function (s) { return s.name; });

    if ((data.headers || []).length) {
      form.querySelectorAll('select[data-source="headers"]').forEach(function (sel) {
        rebuildSelect(sel, data.headers, suggestions[sel.name]);
      });
    }
    if (sheetNames.length) {
      form.querySelectorAll('select[data-source="sheets"]').forEach(function (sel) {
        rebuildSelect(sel, sheetNames, suggestions[sel.name]);
      });
    }

    var cats = suggestions.categories;
    var counts = (data.annotations || {}).categories || {};
    form.querySelectorAll('input[name="categories"]').forEach(function (box) {
      if (cats) box.checked = cats.indexOf(box.value) !== -1;
      var n = counts[box.value];
      var span = box.parentElement.querySelector(".category-name");
      if (span && typeof n === "number") {
        span.textContent = box.value + " (" + n + (n === 1 ? " clip" : " clips") + ")";
      }
    });

    Object.keys(suggestions).forEach(function (key) {
      if (key === "categories") return;
      var el = form.querySelector('[name="' + key + '"]');
      if (el && el.tagName !== "SELECT" && el.type !== "checkbox") el.value = suggestions[key];
    });

    if (banner) {
      var msgs = data.warnings || [];
      banner.textContent = msgs.join("  •  ");
      banner.hidden = msgs.length === 0;
    }
  }

  function analyze() {
    if (!fileInput.files.length) return;
    fetch(url, { method: "POST", body: payload() })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (data) { if (data) apply(data); })
      .catch(function () { console.warn("analyze request failed"); });
  }

  fileInput.addEventListener("change", analyze);
  form.querySelectorAll("select[data-analyze]").forEach(function (sel) {
    sel.addEventListener("change", analyze);
  });
})();
```

- [ ] **Step 4: Update `web/templates/command.html`**

Replace the `<form>` opening tag and the `{% for f in spec.fields %}`
block, and swap the trailing `{% if spec.slug == "sort" %}<script>...`
for the shared script. The full new `{% block content %}` form region:

```jinja
<form method="post" action="/commands/{{ spec.slug }}" enctype="multipart/form-data"
      data-analyze-url="/commands/{{ spec.slug }}/analyze">
  <fieldset>
    <legend>Options</legend>
    <div id="analyze-warnings" class="warn-banner" hidden></div>
    {% for f in spec.fields %}
    {% if f.name == "categories" and spec.slug == "sort" %}
    <fieldset class="categories">
      <legend>{{ f.label }}</legend>
      {% set selected = (values or {}).get("categories") %}
      {% for category in all_categories %}
      <label class="category-checkbox">
        <input type="checkbox" name="categories" value="{{ category }}" {% if selected is none or category in selected %}checked{% endif %}>
        <span class="category-name">{{ category }}</span>
      </label>
      {% endfor %}
      <p class="hint">"3rd parties" is always included — clips from any turned-off category land there for manual review.</p>
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
    <label>
      {{ f.label }}{% if f.required %} (required){% endif %}
      {% if f.type == "checkbox" %}
        <input type="checkbox" name="{{ f.name }}" {% if (values or {}).get(f.name, f.default) %}checked{% endif %}>
      {% else %}
        <input type="text" name="{{ f.name }}" value="{{ (values or {}).get(f.name, f.default) if (values or {}).get(f.name, f.default) is not none else '' }}" {% if f.required %}required{% endif %}>
      {% endif %}
    </label>
    {% endif %}
    {% endfor %}
  </fieldset>
  <label>File(s)
    <input type="file" name="files" multiple required accept=".xlsx">
  </label>
  <button type="submit" name="intent" value="run">Run {{ spec.title }}</button>
  <input type="text" name="preset_name" placeholder="Preset name">
  <label><input type="checkbox" name="overwrite"> Overwrite</label>
  <button type="submit" name="intent" value="save_preset" formaction="/presets/{{ spec.slug }}" formnovalidate>Save as preset</button>
</form>
<script src="/static/analyze.js"></script>
```

Delete the old `{% if spec.slug == "sort" %}<script> ... </script>{% endif %}`
block entirely (the shared `analyze.js` replaces it).

- [ ] **Step 5: Add styles to `web/static/style.css`**

Append:

```css
.warn-banner { background: #fff3cd; border: 1px solid #ffca2c; padding: 0.5rem 0.75rem; margin-bottom: 0.75rem; font-size: 0.9em; }
select { font: inherit; margin-top: 0.25rem; }
```

- [ ] **Step 6: Run the form tests**

Run: `python -m pytest tests/test_web_command_form.py -v`
Expected: PASS (5 tests)

- [ ] **Step 7: Run the full suite**

Run: `python -m pytest -q`
Expected: PASS (~420 tests). If `tests/test_web_final_fixes.py` or
`tests/test_web_presets_ui.py` assert on the old `<input name="name_column">`
text field, update those assertions to expect `<select name="name_column"`
and keep their intent.

- [ ] **Step 8: Commit**

```bash
git add web/static/analyze.js web/templates/command.html web/static/style.css tests/test_web_command_form.py
git add tests/test_web_final_fixes.py tests/test_web_presets_ui.py 2>/dev/null || true
git commit -m "$(cat <<'EOF'
feat(web): analyze-and-suggest wizard on the command form

File pick posts to /commands/{slug}/analyze; column and sheet fields
become dropdowns populated from the file, Sort's category boxes get
live counts, analyze warnings show in a banner. One generic script for
every command; the Sort-only inline script is gone.

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

### Task 7: End-to-end verification against a realistic file

**Files:**
- Test: `tests/test_web_wizard_e2e.py` (create)

**Interfaces:**
- Consumes: everything above.
- Produces: one black-box test proving analyze → edit → run works through
  the HTTP layer with a multi-sheet, multi-category workbook.

- [ ] **Step 1: Write the end-to-end test**

Create `tests/test_web_wizard_e2e.py`:

```python
import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from web.auth import hash_password
from web.main import app


def _client(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    c = TestClient(app)
    c.post("/login", data={"password": "pw"})
    return c


def _master():
    wb = Workbook()
    ws = wb.active
    ws.title = "Master"
    ws.append(["Media File", "Clip Duration"])
    for name in [
        "BM1234_liberation.mxf",                 # AP
        "BM5678_parade.mxf",                     # AP
        "045_SO_EP18_01_3DExplainer_TXLS.mov",   # GFX
        "shutterstock_777.mp4",                  # Shutterstock
        "random_unclassified_xyz.mov",           # 3rd parties
    ]:
        ws.append([name, "00:00:10:00"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_analyze_then_run_sort(monkeypatch):
    client = _client(monkeypatch)
    content = _master()

    # 1. Analyze with the real (non-default) name column.
    analyze = client.post(
        "/commands/sort/analyze",
        data={"name_column": "Media File"},
        files={"files": ("master.xlsx", content, "application/octet-stream")},
    ).json()
    assert analyze["headers"] == ["Media File", "Clip Duration"]
    assert analyze["annotations"]["categories"]["AP"] == 2
    assert set(analyze["suggestions"]["categories"]) >= {"AP", "GFX", "Shutterstock"}
    assert not analyze["warnings"]

    # 2. User keeps only AP + GFX, then runs.
    run = client.post(
        "/commands/sort",
        data={"name_column": "Media File", "categories": ["AP", "GFX"]},
        files={"files": ("master.xlsx", content, "application/octet-stream")},
    )
    assert run.status_code == 200
    wb = load_workbook(io.BytesIO(run.content))
    assert "AP" in wb.sheetnames and "GFX" in wb.sheetnames
    assert "Shutterstock" not in wb.sheetnames
    ap = wb["AP"]
    assert ap.max_row == 3  # header + 2 AP clips
```

- [ ] **Step 2: Run it**

Run: `python -m pytest tests/test_web_wizard_e2e.py -v`
Expected: PASS

- [ ] **Step 3: Manual smoke test**

```bash
BIGMEDIA_WEB_PASSWORD_HASH=$(python -c "from web.auth import hash_password; print(hash_password('dev'))") \
  python -m uvicorn web.main:app --port 8000
```

In a browser: log in with `dev`, open `/commands/sort`, pick a real
`.xlsx`. Confirm the "Filename column" dropdown fills with the file's
headers, category checkboxes show `(N clips)` and auto-tick, changing the
column dropdown re-runs analysis, and Run still downloads a sorted
workbook. Stop the server.

- [ ] **Step 4: Full suite + commit**

Run: `python -m pytest -q`
Expected: PASS (all)

```bash
git add tests/test_web_wizard_e2e.py
git commit -m "$(cat <<'EOF'
test: end-to-end analyze-then-run coverage for the Sort wizard

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 5: Update the web design doc**

In `docs/superpowers/specs/2026-09-01-web-ui-design.md`, add a short
"Analyze step" subsection (2-4 sentences) pointing at
`2026-09-10-analyze-suggest-wizard-design.md` as the detail, so the
original design doc doesn't silently drift. Commit:

```bash
git add docs/superpowers/specs/2026-09-01-web-ui-design.md
git commit -m "$(cat <<'EOF'
docs: note the analyze step in the web UI design doc

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
EOF
)"
```

---

## Self-Review

**1. Spec coverage:**

| Spec item | Task |
|---|---|
| `analyze_files()` — header union, sheets+rows, warnings | Task 1 |
| `analyze_sort()` — wraps `count_categories`, warning not raise | Task 2 |
| `FieldSpec.options_source`, `CommandSpec.analyze`, wiring | Task 3 |
| `web/analyze.py` `run_analysis()` merge + signature-filtered kwargs | Task 4 |
| Delete `sort_analyze`, add generic `/commands/{slug}/analyze`, JSON shape | Task 5 |
| `<select>` rendering, generic `analyze.js`, warning banner, CSS | Task 6 |
| Preset-vs-suggestion (analyze wins) | Task 6 (JS `apply` overwrites inputs/checkboxes after preset render) + Task 7 assertion |
| Multi-file union + disagreement warning | Task 1 (logic) + Task 5 (`test_analyze_union_across_multiple_files`) |
| Shaky analysis → warning, Run stays enabled | Task 2, Task 5 (`test_analyze_missing_name_column_warns_but_200`), Task 6 (banner, no disable) |
| Non-xlsx / corrupt → 400 on analyze | Task 5 (`test_analyze_rejects_non_xlsx`; `except Exception` path) |
| Edge: no file picked → no analyze call | Task 6 (`analyze()` early-returns on empty `files`) |
| Edge: single-sheet file → picker still populates | Task 1 covers 1-sheet output; JS populates from any non-empty list |
| Out of scope: other commands' analyzers, upload caching | Not implemented; framework left ready (Task 3 marks fu-grid's `sheet` field; group's multi-value `sheets` deferred) |

No uncovered spec requirements.

**2. Placeholder scan:** No "TBD"/"handle edge cases"/"similar to"/bare
"write tests" — every code and test step has literal content.

**3. Type consistency:**
- `analyze_files(paths) -> {"headers", "sheets", "warnings"}` — same keys
  in Task 1 definition, Task 4 consumption, Task 5 route.
- `analyze_sort(paths, name_column=...)` returns `suggestions`/
  `annotations`/`warnings` — Task 2 defines, Task 4 merges those exact
  keys.
- `run_analysis(spec, paths, form)` — Task 4 signature; Task 5 calls it
  as `run_analysis(spec, [str(p) for p in paths], form)`. Match.
- `options_source` attribute name — Task 3 defines, Task 6 template reads
  `f.options_source`. Match.
- `data-analyze-url` / `data-source` / `data-analyze` / `#analyze-warnings`
  / `.warn-banner` / `.category-name` — identical strings in Task 6's
  `analyze.js`, template, CSS, and Task 6 tests.
- JSON keys `headers`/`sheets`/`suggestions`/`annotations`/`warnings` and
  `sheets[].name`/`sheets[].rows` — consistent Task 4 → Task 5 tests →
  Task 6 JS → Task 7.

---

## Execution Handoff

Plan complete. Two execution options:

1. **Subagent-Driven (recommended)** — fresh subagent per task, review between tasks.
2. **Inline Execution** — tasks in this session with checkpoints.
