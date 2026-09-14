# Sheet-aware analyze suggestions for group / fu-grid / getty-ids

**Date:** 2026-09-14
**Status:** Approved, ready for implementation plan
**Follows up:** [2026-09-10-analyze-suggest-wizard-design.md](2026-09-10-analyze-suggest-wizard-design.md)

## Problem

The analyze-and-suggest wizard (previous spec) reads header suggestions from
each uploaded workbook's **active sheet** — a deliberate simplification that
matches `sort`/`dedupe`, which really do process the active sheet. It does
not match `group`, `fu-grid`, or `getty-ids`, which hunt for specific named
sheets (`Getty Videos`, `3rd parties`, ...) regardless of which tab Excel
last had open. A real file's active sheet is whatever a person was last
looking at when they saved it — e.g. a "COST" budget tab — which has nothing
to do with the clip-list sheet the command actually reads.

Symptom seen live: uploading two Getty-IDs episode files whose active tab is
"COST" (containing just `"EP1"`/`"EP2"` in cell A1) makes the Filename
Column dropdown offer `EP1`/`EP2` as if they were headers, and raises a
bogus "columns differ" warning comparing COST-tab contents between files.

Goal: `group`, `fu-grid`, and `getty-ids` get header/column suggestions
scoped to the sheet(s) they actually read, plus a sheet-picker for every
sheet selection those commands make (video sheet, stills sheet, the set of
category sheets to group) — populated from the upload's real sheet names,
not typed by hand.

## Scope

In scope: `group`, `fu-grid`, `getty-ids`.
Out of scope (explicitly, per user decision): `sort`, `dedupe`, `compare`,
`fix-getty`. `sort`/`dedupe` are correct as-is (single active-sheet raw
dumps, no sheet field needed). `compare`/`fix-getty` use the separate
pair-upload page (`pair_command.html`), which has no analyze/JS wiring at
all today — wiring it up is a separate, larger piece of work left for later.

## Decisions (from brainstorming)

1. Fix is scoped to the three multi-file commands that read named sheets;
   pair-mode commands untouched.
2. `group`'s free-text "Sheets to group" field becomes a checkbox list built
   from the upload's real sheet names, pre-checked by matching
   `DEFAULT_SHEETS`/`SHEET_ALIASES` — same pattern as Sort's category
   checkboxes, but built dynamically from upload content instead of a fixed
   server-rendered list.
3. Header/duration suggestions use the **union** of headers across whichever
   sheet(s) are currently selected for that command (not the intersection) —
   consistent with the existing tolerant "not found → substitute + warn" UX,
   and with the fact real files already disagree on header naming
   sheet-to-sheet (that's what the alias tables are for).
4. A stills-sheet selection is allowed to resolve to nothing — "this project
   has no stills" is a legitimate, common state (confirmed on the Wild
   Return project), not an error to warn about or paper over with a wrong
   guess.

## Architecture

No new endpoint, no new request/response shape at the transport level — the
existing `/commands/{slug}/analyze` route and JSON contract from the
previous spec are reused. The change is entirely in what `headers` (and,
for `group`, a new checkbox list) resolve to, driven by a new declarative
`sheet_source` on `FieldSpec` that `web/analyze.py` resolves using the
*current* form values already available to it (same mechanism
`_analyzer_kwargs` uses today).

```
FieldSpec("name_column", ..., options_source="headers", sheet_source=["sheet_name", "stills_sheet_name"])
                                                              │
run_analysis(spec, paths, form)                              │ reads current
   ├── xlsx_utils.analyze_files(paths)                       │ value(s) of
   │      → headers (active-sheet, unchanged)                │ those field(s)
   │      → headers_by_sheet: {sheet_name: [headers]}  ◀──────┘ from `form`
   │      → sheets, warnings (unchanged)
   └── if any field declares sheet_source:
          result["headers"] = union of headers_by_sheet[resolved sheet]
                               for every resolved sheet, case-insensitively
                               matched against the file's real sheet names
          result["warnings"] = per-resolved-sheet disagreement warnings
                               (replaces the active-sheet-based warning for
                               this command's response only)
```

Because every `options_source="headers"` field within a single command
already shares one flat `headers` list client-side (existing JS, unchanged),
and because every headers field in `group`/`fu-grid`/`getty-ids` scopes to
the *same* sheet selection within that command, one resolved `headers` list
per analyze call is sufficient — no per-field header lists, no JS changes
for this part.

## Components

### `bigmedia/xlsx_utils.py` — `analyze_files(paths)`

Extend, don't replace. Add a `headers_by_sheet` computation alongside the
existing active-sheet `headers`:

```python
{
    "headers": [...],            # unchanged: active-sheet union (sort/dedupe)
    "headers_by_sheet": {        # new: every sheet, every file
        "Getty Videos": ["Track", "Clip Name", "Seconds"],
        "Getty Stills": [...],
        "COST": [...],
        ...
    },
    "sheets": [...],              # unchanged
    "warnings": [...],            # unchanged (active-sheet based; see below)
}
```

- Keys are the **first-seen casing** of each sheet title; lookups elsewhere
  are case-insensitive against this dict (reuses the matching convention
  already in `getty_ids._find_sheet` / `group_duplicates.py`'s
  `title.strip().lower()` checks).
- Per-sheet header list is the union across files that have a sheet by that
  name (same first-seen-order union rule as the existing `headers` field).
- No new warnings here — sheet-scoped disagreement warnings are computed in
  `web/analyze.py`, which knows which sheets are actually relevant to the
  current command; `analyze_files` stays command-agnostic.
- Sheet iteration replaces the single `wb.active` read with a loop over
  `wb.sheetnames`; still `read_only=True, data_only=True`, still one pass
  per file (only row 1 of each sheet is read — cheap).

Unit-tested in `tests/test_xlsx_utils.py` (extends existing
`analyze_files` tests).

### New shared helper — `bigmedia/xlsx_utils.py` — `match_sheet_name(sheet_names, target, aliases=())`

Small pure function: case-insensitive match of `target` against
`sheet_names`, falling back through `aliases` (each an alternate name to
try), returning the real sheet name (original casing) or `None`. Formalizes
the case-insensitive-match-plus-alias-fallback pattern that already exists
ad hoc in `getty_ids.py` (`_find_sheet`/`_find_sheet_any`) and
`group_duplicates.py` (inline `.strip().lower()` checks). Both `web/analyze.py`
(new suggestion logic) and, opportunistically, `getty_ids.py`'s existing
`_find_sheet`/`_find_sheet_any` are rewritten in terms of it, so the web
suggestion layer and the actual report-building logic can never disagree
about "does this file have a stills sheet". `group_duplicates.py`'s inline
matching is left as-is (different shape — it matches a *set* of wanted
sheets against every sheet in one file, not one target against a name
list) — not touched, out of scope.

Unit-tested alongside the existing `extract_getty_id` etc. tests in
`tests/test_xlsx_utils.py`.

### `web/commands.py` — `FieldSpec` gains `sheet_source`

```python
@dataclass
class FieldSpec:
    name: str
    label: str
    type: str
    default: Any = None
    required: bool = False
    options_source: str = None       # None | "headers" | "sheets"
    sheet_source: list = None        # names of this command's sheet field(s)
                                      # that scope a "headers" field's options
    allow_missing_sheet: bool = False     # "sheets" fields only: True allows a
                                      # blank "(none)" option (stills_sheet_name)
```

Per-command wiring (replacing the shared `_NAME_COLUMN`/`_DURATION_COLUMN`
constants for these three commands only — they now need distinct
`sheet_source` per command, so each gets its own local `FieldSpec`
instance instead of sharing one; `sort`/`dedupe`/`compare` keep using the
shared constants unchanged):

- **fu-grid**: `sheet` (already `options_source="sheets"`, default
  `"3rd parties"`) unchanged in shape. `name_column`/`duration_column` gain
  `sheet_source=["sheet"]`.
- **getty-ids**: `sheet_name` and `stills_sheet_name` change from plain
  text to `options_source="sheets"` (video: default `allow_missing_sheet=
  False`; stills: `allow_missing_sheet=True`). `name_column` gains `sheet_source=["sheet_name",
  "stills_sheet_name"]`. `seconds_column` (currently plain text) also
  becomes `options_source="headers"` with the same `sheet_source` — it's a
  real column pickable off the same sheets, no reason it stays free text
  while `name_column` gets a dropdown.
- **group**: `sheets` changes from `type="list"` free text to a new
  `type="sheet_checklist"`, `options_source="sheets"`. `name_column`/
  `duration_column` gain `sheet_source=["sheets"]` (the checklist's current
  checked values, plural).

### `web/analyze.py` — resolving `sheet_source`

`run_analysis` gains, after the existing generic `analyze_files` call and
before the per-command `spec.analyze` hook:

```python
sheet_fields = {f.name for f in spec.fields if f.options_source == "sheets"}
headers_fields = [f for f in spec.fields if f.sheet_source]

if headers_fields:
    resolved_sheets = set()
    for f in headers_fields:
        for source_name in f.sheet_source:
            value = _form_value(form, fields_by_name[source_name])
            for name in (value if isinstance(value, list) else [value]):
                if name:
                    resolved_sheets.add(name)

    matched = {
        name: match_sheet_name(sheet_names_in_files, name, aliases=SHEET_ALIASES.get(name.lower(), ()))
        for name in resolved_sheets
    }
    result["headers"] = union of headers_by_sheet[m] for m in matched.values() if m
    result["warnings"] += [
        f"{name}: not found in {file}" or per-sheet header-disagreement warnings
        ...
    ]
```

Warning text: `f"{file_name}: '{sheet}' sheet columns differ from {first_file_with_that_sheet}"`
for a header disagreement, `f"{file_name}: '{sheet}' sheet not found"` for a
field with `allow_missing_sheet=False` whose resolved sheet is missing from
a file entirely.
Never warn about a missing *optional* sheet — that's the "no stills" case,
expected and silent.

Suggestions for the new `sheets`-sourced fields (`sheet_name`,
`stills_sheet_name`, `group`'s checklist) are produced the same way
Sort's `categories` suggestion is today — computed here in `run_analysis`
using `match_sheet_name` against `DEFAULT_SHEETS`/`SHEET_ALIASES` (group) or
the field's own default name (`"Getty Videos"`/`"Getty Stills"`/
`"3rd parties"`) — no per-command `spec.analyze` hook needed for this, since
it's the same shape across all three commands.

### `web/templates/command.html` — `sheet_checklist` field type

New branch alongside the existing `categories`-special-case, generic (not
hardcoded to `group`):

```html
{% elif f.type == "sheet_checklist" %}
<fieldset class="sheet-checklist" data-analyze data-source="sheets" data-checklist-name="{{ f.name }}">
  <legend>{{ f.label }}</legend>
  <p class="hint">Pick a file to see its sheets.</p>
</fieldset>
```

Empty at render time (no sheet names known before upload) — JS populates
it, mirroring `rebuildSelect` but emitting `<label><input type=checkbox
name="{{ f.name }}" value="...">...</label>` per real sheet name instead of
`<option>`s. Same "preserve valid current selection, otherwise fall back to
the server suggestion" rule as `rebuildSelect`.

A hidden `sheets_present` marker (same idea as Sort's `categories_present`)
distinguishes "every box unchecked on purpose" from "field omitted" so
`command_submit` can pass an explicit empty list rather than falling back to
`group_duplicates_workbook`'s `sheets or DEFAULT_SHEETS` default.

### `web/static/analyze.js`

- New `rebuildChecklist(fieldset, options, suggested, name)`, parallel to
  `rebuildSelect`: builds one checkbox per option, preserves any
  currently-checked box whose value is still in `options`, otherwise checks
  exactly the boxes in `suggested`. Wired into `apply()` next to the
  existing `select[data-source="sheets"]` handling, driven by
  `fieldset[data-checklist-name]` elements.
- `rebuildSelect` rebuilds a select's options from scratch on every call
  (`sel.innerHTML = ""` then repopulate from `options`), so a template-only
  "(none)" option would be wiped by the first analyze response — the
  distinction has to travel through `apply()`. `command.html` marks an
  optional `sheets`-select with `data-optional="true"` (from
  `f.allow_missing_sheet is True`); `apply()` passes
  `sel.dataset.optional === "true" ? [""].concat(sheetNames) : sheetNames`
  as the `options` array for that one select, with a small label map
  (`{"": "(none)"}`) so `rebuildSelect`'s generated `<option>` shows
  "(none)" instead of an empty string. Required `sheets`-selects
  (`sheet_name`, fu-grid's `sheet`) are untouched — no blank option, so an
  unresolved required sheet still falls back to a real (if wrong) sheet
  name, which is visibly wrong and promptable rather than silently blank.
- No change to how `analyze()` triggers re-analysis on a checked/unchecked
  checkbox — it already listens for `change` on `select[data-analyze]`;
  extend the same listener registration to also cover
  `fieldset[data-checklist-name] input[type=checkbox]`, including ones
  created after the initial page load (event delegation on the `fieldset`,
  not per-checkbox listeners).

## JSON contract (delta from the previous spec)

```json
{
  "headers": ["Clip Name", "Seconds"],
  "headers_by_sheet": {"Getty Videos": ["Track", "Clip Name", "Seconds"], "Getty Stills": ["Track", "Clip Name", "Seconds"], "COST": ["EP1"]},
  "sheets": [{"name": "Getty Videos", "rows": 39}, {"name": "COST", "rows": 4}],
  "suggestions": {"name_column": "Clip Name", "sheet_name": "Getty Videos", "stills_sheet_name": ""},
  "warnings": []
}
```

`headers_by_sheet` is new; everything else keeps its existing shape.
`headers` for `group`/`fu-grid`/`getty-ids` is now sheet-scoped instead of
active-sheet; unchanged for `sort`/`dedupe`.

## Testing

| File | Adds |
|---|---|
| `tests/test_xlsx_utils.py` | `analyze_files`: `headers_by_sheet` per sheet per file, union across files with the same sheet name; `match_sheet_name`: case-insensitive, alias fallback, no match → `None` |
| `tests/test_web_routes_getty_ids.py` | analyze response: `name_column`/`seconds_column` headers scoped to `sheet_name`+`stills_sheet_name`, not the active COST-style sheet; `sheet_name`/`stills_sheet_name` suggestions; missing stills sheet → no warning |
| new `tests/test_web_routes_group_analyze.py` (or extend existing group tests) | `sheets` checklist suggestion pre-checks `DEFAULT_SHEETS`/alias matches only; `name_column`/`duration_column` headers scoped to checked sheets |
| new `tests/test_web_routes_fu_grid_analyze.py` (or extend existing) | `name_column`/`duration_column` headers scoped to the selected `sheet` |
| `tests/test_web_commands.py` | `command.html` renders the `sheet_checklist` fieldset for `group`; existing `options_source="sheets"` rendering now also covers `getty-ids`' `sheet_name`/`stills_sheet_name` |

Full `pytest` green (currently 444 passing).

## Edge cases

- Stills sheet absent from every uploaded file → `stills_sheet_name`
  suggests `""` ("(none)"), no warning, `seconds_column`/`name_column`
  headers come from the video sheet alone.
- Video sheet name typed/selected doesn't match anything in the upload →
  warning (video's `allow_missing_sheet` defaults to `False`), `headers`
  falls back to whatever *did* resolve (stills, if any) rather than going
  empty.
- `group`: user unchecks every checkbox → `sheets_present` marker present,
  empty list → explicit "group nothing", output is a straight pass-through
  copy (this is `group_duplicates_workbook`'s existing behavior for an
  empty `sheets` list, unchanged — just now reachable from the UI).
- Two uploaded files whose "Getty Videos" sheets have different headers but
  whose active/COST sheets happen to match → warning now correctly fires
  (previously silent, since only the active sheet was compared).
- `sort`/`dedupe` behavior is byte-for-byte unchanged — no `sheet_source`
  declared, `run_analysis` takes the untouched `headers` path.

## Follow-ups (not this spec)

- `compare`/`fix-getty` (pair-mode) get no analyze/suggest wiring at all
  yet — a separate spec, since it needs a second analyze endpoint shape
  (two named files, not a file list) and new template/JS for
  `pair_command.html`.
- Unifying `getty_ids._STILLS_SHEET_ALIASES` and
  `group_duplicates.SHEET_ALIASES` (same fact, two tables) — noted, not
  fixed here; `match_sheet_name` takes `aliases` as a parameter so either
  table can feed it without forcing that unification now.
