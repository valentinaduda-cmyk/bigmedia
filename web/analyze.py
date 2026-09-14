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
            targets = _resolve_sheet_field_values(form, source_field)
            if not targets and source_field.type == "sheet_checklist" and not form.get(f"{source_field.name}_present"):
                # Checklist hasn't been populated by the client yet (this is
                # the first analyze response before any rebuild) -- resolve
                # against its own suggested sheets instead of resolving
                # nothing, so headers/suggestions aren't empty on the very
                # first response. Once the client has populated the
                # checklist at least once, its "_present" marker is always
                # sent from then on, and an explicitly empty selection is
                # honored as a real "user unchecked everything."
                targets = _suggest_sheet_value(source_field, generic, form) or []
            for target in targets:
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


def _suggest_sheet_value(field, generic, form):
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

    current = _form_value(form, field)
    if isinstance(current, list):
        current = current[0] if current else None
    if current in real_sheet_names or (current == "" and field.allow_missing_sheet):
        return None  # already a valid or deliberate selection -- don't override it

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
            suggestion = _suggest_sheet_value(field, generic, form)
            if suggestion is not None:
                result["suggestions"].setdefault(field.name, suggestion)

    if spec.analyze:
        extra = spec.analyze(paths, **_analyzer_kwargs(spec, form))
        result["suggestions"].update(extra.get("suggestions", {}))
        result["annotations"].update(extra.get("annotations", {}))
        result["warnings"].extend(extra.get("warnings", []))
    return result
