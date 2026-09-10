# Analyze-and-suggest wizard for web commands

**Date:** 2026-09-10
**Status:** Approved, ready for implementation plan

## Problem

The web UI's command forms are blank on load. The user uploads a file and
must know, unprompted, which header holds the clip name, which categories
are worth keeping, which sheets exist. Sort has a partial fix — JS posts to
`/commands/sort/analyze` on file pick and ticks category checkboxes from a
live count — but it is Sort-only, single-purpose, and has no explicit
review step.

Goal: every command follows one flow — **upload → analyze → suggest →
user confirms/edits → run**. The app inspects the uploaded file(s), fills
the blanks with suggestions, the user eyeballs and adjusts, then runs.

## Scope

Build the generic wizard framework and convert **Sort** first. The other
six commands (dedupe, group, fu-grid, compare, fix-getty, getty-ids) are
follow-ups — the framework must support them, but wiring each is out of
scope here.

Also out of scope: server-side upload caching, progress indicators.

## Decisions (from brainstorming)

1. **Rollout:** framework + Sort now; other commands one at a time later.
2. **File between steps:** browser keeps the file in the `<input>` and
   re-sends it with the final Run. No server-side upload storage, no
   cleanup, server stays stateless. Files are KB–low MB; double transfer
   is cheap.
3. **Wizard shape:** one page, progressive reveal. File pick triggers
   analyze; the options section is populated in place. No new GET pages.
4. **Filename column:** auto-detected from known aliases, shown as a
   `<select>` of all detected headers with the guess pre-selected.
   Changing it re-fires analyze. No separate "now analyze" button.
5. **Multiple files (batch commands):** union — headers/sheets/counts
   merged across all uploaded files, one set of options applied to all,
   with a warning line when files disagree.
6. **Shaky analysis** (unreadable, no name column, missing sheet):
   warning banner in the options area, options fall back to defaults,
   Run stays enabled. Hard errors (not `.xlsx`, corrupt) still hard-fail
   on Run as today.
7. **Presets vs suggestions:** analyze wins. File-specific suggestions
   overwrite a loaded preset's values for fields analyze has an opinion
   on (name column, categories); the preset still sets fields analyze
   does not touch (autofit, widths, fonts).

## Architecture

Generic analysis for every command, plus an optional per-command hook.

```
browser (file change / dropdown change)
   │  POST /commands/{slug}/analyze   (file[s] + current field values)
   ▼
web/main.py  generic route
   │
   ▼
web/analyze.py  run_analysis(spec, paths, field_values)
   ├── bigmedia/xlsx_utils.analyze_files(paths)     → headers, sheets, warnings
   └── spec.analyze(paths, field_values)  (optional) → command-specific extras
   │
   ▼
JSON  { headers, sheets, suggestions, annotations, warnings }
   │
   ▼
browser  apply suggestions to inputs, rebuild <select> options,
         paint annotations, fill warning banner
```

### JSON contract

```json
{
  "headers": ["Clip Name", "Clip Duration", "Source Reel Name"],
  "sheets": [{"name": "Sheet1", "rows": 412}],
  "suggestions": { "name_column": "Clip Name", "categories": ["AP", "Getty Videos"] },
  "annotations": { "categories": {"AP": 12, "Getty Videos": 3, "Reuters": 0} },
  "warnings": ["master_ep3.xlsx: no 'Clip Name' column — using 'Name'"]
}
```

- `headers` — union across files, first-seen order. Feeds every column
  dropdown.
- `sheets` — union, with row counts. Feeds sheet-picker fields.
- `suggestions` — new values for named fields, applied on top of
  preset/defaults (decision 7).
- `annotations` — display-only decoration per field (count suffix on
  Sort's category checkboxes).
- `warnings` — shown in a banner in the options area; never blocks Run.

Any key may be absent/empty. The browser no-ops on missing keys.

## Components

### `bigmedia/xlsx_utils.py` — `analyze_files(paths)`

New pure read-only function. Opens each workbook `read_only=True,
data_only=True`. Returns:

```python
{
    "headers": [...],   # union, first-seen order, row-1 values, blanks skipped
    "sheets": [{"name": str, "rows": int}, ...],   # union; rows = max_row - 1, floored at 0
    "warnings": [...],   # e.g. a file whose header set differs from the first file's
}
```

- Header union is over every sheet's row 1? No — over each workbook's
  **active** sheet row 1 (matches how batch commands read input).
- A workbook that fails to open → one warning `"<name>: could not read
  ({exc})"`, that file skipped, others still processed.
- Disagreement warning: if a later file's active-sheet header set is not
  equal to the first file's, emit `"<name>: columns differ from
  <first name>"`.

Unit-tested in `tests/test_xlsx_utils.py`.

### `bigmedia/sort_workbook.py` — `analyze_sort(paths, name_column)`

New function, wraps the existing `count_categories`:

```python
def analyze_sort(paths, name_column="Clip Name"):
    try:
        counts = count_categories(paths, name_column=name_column)
    except ValueError as exc:
        return {"warnings": [str(exc)]}
    return {
        "suggestions": {"categories": [c for c, n in counts.items() if n > 0]},
        "annotations": {"categories": counts},
    }
```

`count_categories` already handles the name-column aliases and
`read_only`. Unit-tested in `tests/test_sort_workbook.py`.

### `web/commands.py`

- `FieldSpec` gains `options_source: str | None = None`. Value
  `"headers"` or `"sheets"` marks the field as an analyze-populated
  `<select>`.
- `_NAME_COLUMN` and `_DURATION_COLUMN` gain `options_source="headers"`.
  (Group's `sheets` list field and fu-grid's `sheet` field gain
  `options_source="sheets"` — rendered generically now, exercised when
  those commands get their analyzers.)
- `CommandSpec` gains `analyze: Callable | None = None`.
- Sort's spec gains `analyze=analyze_sort`.

### `web/analyze.py` (new)

```python
def run_analysis(spec, paths, field_values) -> dict:
    result = {"headers": [], "sheets": [], "suggestions": {},
              "annotations": {}, "warnings": []}
    generic = analyze_files(paths)
    result["headers"] = generic["headers"]
    result["sheets"] = generic["sheets"]
    result["warnings"] += generic["warnings"]
    if spec.analyze:
        extra = spec.analyze(paths, **_analyze_kwargs(spec, field_values))
        result["suggestions"].update(extra.get("suggestions", {}))
        result["annotations"].update(extra.get("annotations", {}))
        result["warnings"] += extra.get("warnings", [])
    return result
```

`_analyze_kwargs` passes only the field values the analyzer signature
needs (Sort: `name_column`).

### `web/main.py`

- Delete `sort_analyze`.
- Add:

```python
@app.post("/commands/{slug}/analyze", dependencies=[Depends(require_login)])
async def command_analyze(request: Request, slug: str,
                          files: List[UploadFile] = File(...)):
    spec = _command_or_404(slug)
    bad = reject_non_xlsx([f.filename for f in files])
    if bad:
        return JSONResponse({"error": f"Not an .xlsx file: {', '.join(bad)}"},
                            status_code=400)
    form = await request.form()
    with tempfile.TemporaryDirectory() as tmp:
        paths = _save_uploads(files, Path(tmp))
        try:
            data = run_analysis(spec, [str(p) for p in paths], form)
        except Exception as exc:
            return JSONResponse({"error": str(exc)}, status_code=400)
    return JSONResponse(data)
```

- The Run path (`command_submit`, `_run_batch`, `_run_combine`,
  `pair_submit`) is unchanged — the final POST still carries every field
  plus the files.

### `web/templates/command.html`

- Fields with `options_source` render as `<select name=...>` seeded with
  a single `<option>` for the current value; JS adds the rest after
  analyze.
- Replace the Sort-only `<script>` with one generic script (in
  `base.html` or a small `web/static/analyze.js`) that:
  - on `change` of `input[type=file]` and of any `select[data-analyze]`,
    POSTs the file(s) + current field values to
    `/commands/{{ spec.slug }}/analyze`;
  - rebuilds `select[data-analyze][data-source=headers]` options from
    `headers`, `select[...][data-source=sheets]` from `sheets`
    (preserving current selection when still valid);
  - applies `suggestions` — sets input values, checks/unchecks category
    boxes, selects dropdown options;
  - paints `annotations` — e.g. `"AP (12 clips)"` next to a checkbox;
  - fills `#analyze-warnings` banner, empty when no warnings;
  - on fetch failure: leaves inputs at defaults, one `console.warn`.
- Keep the category-checkbox `<fieldset>` block; the generic script
  drives it via `annotations.categories` + `suggestions.categories`.

### `web/static/style.css`

Minor: `#analyze-warnings` banner style (amber, distinct from the red
`.error-banner`), `select` sizing consistent with text inputs.

## Testing

| File | Adds |
|---|---|
| `tests/test_xlsx_utils.py` | `analyze_files`: header union, sheet+row counts, multi-file disagreement warning, unreadable file → warning + others still processed |
| `tests/test_sort_workbook.py` | `analyze_sort`: counts, `suggestions.categories` non-zero only, missing name column → warning not exception |
| `tests/test_web_routes_sort_analyze.py` | reworked: generic `/commands/{slug}/analyze` returns `headers`/`sheets`/`suggestions`/`annotations`; login gate (303); non-xlsx 400; multi-file union; keep existing Run-path assertions |
| `tests/test_web_commands.py` | `command.html` renders `<select>` for `options_source` fields; generic analyze script present for a non-sort command |

Full `pytest` green (currently 394 passing).

## Edge cases

- No file picked yet → dropdowns show default value only, no analyze call.
- Name column absent from every file → `count_categories` falls through
  aliases; still nothing → warning, categories at defaults, Run enabled.
- Single-sheet file → sheet pickers populate with one option.
- Name-column dropdown changed → re-analyze → counts recompute.
- Preset loaded then file picked → suggestions overwrite `name_column` /
  `categories`; preset keeps autofit/width/font.
- `.xls` / corrupt on analyze → 400, banner shows message, no partial
  state applied.

## Follow-ups (not this spec)

- Per-command analyzers for dedupe (dupe-count preview), group
  (name+duration+fps+sheets), fu-grid (source sheet + columns), and the
  pair/combine commands.
- Consider a dupe/row-count preview surface shared across commands.
