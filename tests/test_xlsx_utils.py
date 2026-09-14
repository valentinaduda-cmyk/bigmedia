import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from bigmedia.xlsx_utils import (
    analyze_files, match_sheet_name,
    HEADER_FILL, apply_header_style, sample_data_font, first_data_font,
    style_output_sheets, measure_column_widths,
    AUTOFIT_MIN_WIDTH, AUTOFIT_MAX_WIDTH, AUTOFIT_PADDING,
    is_backup_sheet,
)


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


def test_non_string_header_cells_are_stringified(tmp_path):
    a = _make(tmp_path / "a.xlsx", {"S": [
        [datetime.datetime(2026, 9, 10), 42, "Notes"],
        ["x", 1, "n"],
    ]})
    result = analyze_files([a])
    assert all(isinstance(h, str) for h in result["headers"])
    assert "42" in result["headers"]
    assert any("2026-09-10" in h for h in result["headers"])


def test_lazily_parsed_xml_failure_warns_and_others_still_processed(tmp_path):
    import zipfile

    # Create a valid xlsx, then corrupt it by removing workbook.xml
    # This passes load_workbook but fails when accessing sheets
    good = _make(tmp_path / "good.xlsx", {"S": [["Clip Name"], ["x"]]})
    corrupt = tmp_path / "corrupt.xlsx"

    # Create a minimal valid zip that looks like an xlsx but is missing critical XML
    with zipfile.ZipFile(corrupt, 'w') as zf:
        # Write minimal required structure that passes initial load but fails on access
        zf.writestr('[Content_Types].xml',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '</Types>')
        zf.writestr('_rels/.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>')
        # Missing xl/workbook.xml will cause failure when accessing sheets

    result = analyze_files([str(corrupt), good])
    assert result["headers"] == ["Clip Name"]
    assert any("corrupt.xlsx" in w and "could not read" in w for w in result["warnings"])


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
    # "Clip Name Column" (>10 chars, so it doesn't clamp to the min) sits at
    # column A on s1 and column B on s2. Header-keyed widths follow the name;
    # index-keyed widths would instead make column A wide on BOTH sheets.
    s1 = _sheet(wb, "S1", [["Clip Name Column", "D"], ["x", "y"]])
    s2 = _sheet(wb, "S2", [["D", "Clip Name Column"], ["y", "x"]])
    style_output_sheets([s1, s2], width_by="header", data_font=Font(name="Calibri", size=11))
    wide = len("Clip Name Column") + 2
    assert s1.column_dimensions["A"].width == wide
    assert s2.column_dimensions["B"].width == wide
    assert s1.column_dimensions["A"].width == s2.column_dimensions["B"].width  # keyed by name
    # the narrow "D" column stays at the min on both, and crucially column A of
    # s2 is NOT wide (which is what index-keying would have produced)
    assert s2.column_dimensions["A"].width == AUTOFIT_MIN_WIDTH
    assert s1.column_dimensions["B"].width == AUTOFIT_MIN_WIDTH


def test_style_output_sheets_skips_worksheet_tab():
    wb = Workbook()
    wb.remove(wb.active)
    ws = _sheet(wb, "Worksheet", [["Clip Name"], ["x"]])
    ws["A1"].fill = PatternFill()  # no fill
    style_output_sheets([ws], width_by="index", data_font=Font(name="Calibri", size=11))
    assert ws["A1"].fill.patternType is None  # untouched


def test_measure_column_widths_clamp_and_padding():
    widths = measure_column_widths([
        ["x" * 200, "y" * 20, "z"],
    ])
    assert widths[1] == AUTOFIT_MAX_WIDTH                 # 200 chars -> clamped
    assert widths[2] == 20 + AUTOFIT_PADDING              # mid-range -> len + padding
    assert widths[3] == AUTOFIT_MIN_WIDTH                 # 1 char -> clamped up


def test_apply_header_style_nameless_source_font_gets_fallback():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Clip Name"
    ws["A1"].font = Font()  # no name, no size
    apply_header_style(ws)
    assert ws["A1"].font.name is not None
    assert ws["A1"].font.size is not None
    assert ws["A1"].font.bold is True


def test_is_backup_sheet_matches_worksheet_and_dump_tabs():
    assert is_backup_sheet("Master XML")
    assert is_backup_sheet("EP01 - Master XML")
    assert is_backup_sheet("Kopie listu 1")
    assert is_backup_sheet("Worksheet")
    assert is_backup_sheet("Copy of Sheet1")
    assert not is_backup_sheet("Getty Videos")
    assert not is_backup_sheet("3rd parties")


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
