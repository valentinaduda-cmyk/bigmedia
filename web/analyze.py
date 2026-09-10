import inspect

from bigmedia.group_duplicates import DURATION_COLUMN_ALIASES
from bigmedia.sort_workbook import NAME_COLUMN_ALIASES
from bigmedia.xlsx_utils import analyze_files
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
    Starlette FormData (has getlist) and a plain dict (tests)."""
    if field.type == "list" and hasattr(form, "getlist"):
        return form.getlist(field.name)
    return form.get(field.name)


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
    result = {
        "headers": generic["headers"],
        "sheets": generic["sheets"],
        "suggestions": {},
        "annotations": {},
        "warnings": list(generic["warnings"]),
    }

    # Generic name/duration column guess from known header aliases (spec
    # decision 4). setdefault + running before the per-command analyzer so a
    # command-specific analyzer's suggestions.update() can still override.
    for field, aliases in _COLUMN_SUGGESTIONS:
        match = next((a for a in aliases if a in result["headers"]), None)
        if match:
            result["suggestions"].setdefault(field, match)

    if spec.analyze:
        extra = spec.analyze(paths, **_analyzer_kwargs(spec, form))
        result["suggestions"].update(extra.get("suggestions", {}))
        result["annotations"].update(extra.get("annotations", {}))
        result["warnings"].extend(extra.get("warnings", []))
    return result
