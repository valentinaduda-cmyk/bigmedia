from openpyxl import load_workbook
from bigmedia.sort_workbook import sort_workbook
from bigmedia.classify import classify
from bigmedia.xlsx_utils import find_column


def test_sort_workbook_creates_backup_and_category_sheets(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path))

    wb = load_workbook(out_path)

    # Worksheet backup must be first and match the source, except the
    # "Source Reel Name" column, which is overwritten with each row's
    # resolved category.
    assert wb.sheetnames[0] == "Worksheet"
    src = load_workbook(sample_master).active
    backup = wb["Worksheet"]
    source_col_idx = find_column(src, "Source Reel Name")
    for r in range(1, src.max_row + 1):
        for c in range(1, src.max_column + 1):
            if c == source_col_idx and r > 1:
                continue
            assert backup.cell(row=r, column=c).value == src.cell(row=r, column=c).value

    assert counts["AP"] == 2  # includes the "(1)" duplicate — sort doesn't dedupe
    assert counts["GFX"] == 1
    assert counts["Getty Unknown"] == 1
    assert counts["Getty Videos"] == 1
    assert counts["Shutterstock"] == 1
    assert counts["3rd parties"] == 1

    # Backup rows are labeled with the category they were sorted to.
    name_col_idx = find_column(src, "Clip Name")
    for r in range(2, src.max_row + 1):
        name = src.cell(row=r, column=name_col_idx).value
        if name is None:
            continue
        assert backup.cell(row=r, column=source_col_idx).value == classify(name)


def test_sort_workbook_fills_source_reel_name_per_sheet(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out_path))

    wb = load_workbook(out_path)
    for cat in ["AP", "GFX", "Getty Unknown", "Getty Videos", "Shutterstock", "3rd parties"]:
        ws = wb[cat]
        source_col_idx = find_column(ws, "Source Reel Name")
        for r in range(2, ws.max_row + 1):
            assert ws.cell(row=r, column=source_col_idx).value == cat


def test_sort_workbook_custom_name_column(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path), name_column="Clip Name")
    assert sum(counts.values()) == 7
