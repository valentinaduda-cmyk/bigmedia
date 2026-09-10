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
