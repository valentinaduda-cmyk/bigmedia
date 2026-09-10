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
