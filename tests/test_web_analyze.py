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
