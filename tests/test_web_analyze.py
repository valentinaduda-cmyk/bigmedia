from openpyxl import Workbook

from web.analyze import run_analysis
from web.commands import COMMANDS, CommandSpec, FieldSpec


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
    # Generic column auto-detect (Fix 3) fills name_column; no per-command
    # analyzer runs, so nothing else is suggested and no annotations.
    assert result["suggestions"] == {"name_column": "Clip Name"}
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


# --- Fix 3: name/duration column auto-detect (spec decision 4) --------------

def test_autodetect_name_column_clip_name(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Clip Name", "Clip Duration"], ["x.mov"])
    result = run_analysis(COMMANDS["dedupe"], [path], {})
    assert result["suggestions"]["name_column"] == "Clip Name"


def test_autodetect_name_column_falls_back_to_name(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Name"], ["x.mov"])
    result = run_analysis(COMMANDS["dedupe"], [path], {})
    assert result["suggestions"]["name_column"] == "Name"


def test_autodetect_duration_column(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Clip Name", "Duration"], ["x.mov"])
    result = run_analysis(COMMANDS["dedupe"], [path], {})
    assert result["suggestions"]["duration_column"] == "Duration"


def test_autodetect_absent_when_no_nameish_column(tmp_path):
    path = _sheet(tmp_path / "m.xlsx", ["Media File", "Length"], ["x.mov"])
    result = run_analysis(COMMANDS["dedupe"], [path], {})
    assert "name_column" not in result["suggestions"]
    assert "duration_column" not in result["suggestions"]


def test_per_command_analyzer_still_overrides_autodetect(tmp_path):
    # Sort's analyzer emits suggestions.categories only, not name_column,
    # so the auto-detected name_column must survive alongside it.
    path = _sheet(tmp_path / "m.xlsx", ["Clip Name"], ["BM1234_x.mxf"])
    result = run_analysis(COMMANDS["sort"], [path], {"name_column": "Clip Name"})
    assert result["suggestions"]["name_column"] == "Clip Name"
    assert "categories" in result["suggestions"]


# --- Fix 6: _analyzer_kwargs type-safety + param guards --------------------

def test_analyzer_param_not_in_fields_is_not_passed(tmp_path, monkeypatch):
    path = _sheet(tmp_path / "m.xlsx", ["Clip Name"], ["BM1234_x.mxf"])
    seen = {}

    def fake_analyze(paths, name_column="Clip Name", files=None, bogus=None):
        seen["files"] = files
        seen["bogus"] = bogus
        seen["name_column"] = name_column
        return {"suggestions": {}}

    spec = CommandSpec(
        slug="fake", title="Fake", upload_mode="batch", func=lambda *a, **k: None,
        output_suffix="x", analyze=fake_analyze,
        fields=[FieldSpec("name_column", "Filename column", "text", "Clip Name")],
    )
    run_analysis(spec, [path], {"name_column": "Clip Name", "files": "SHOULD_NOT_LEAK", "bogus": "x"})
    assert seen["files"] is None
    assert seen["bogus"] is None
    assert seen["name_column"] == "Clip Name"
