from openpyxl import load_workbook
from bigmedia.sort_workbook import sort_workbook


def test_sort_workbook_creates_backup_and_category_sheets(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path))

    wb = load_workbook(out_path)

    # Worksheet backup must be first and byte-identical to the source.
    assert wb.sheetnames[0] == "Worksheet"
    src = load_workbook(sample_master).active
    backup = wb["Worksheet"]
    for r in range(1, src.max_row + 1):
        for c in range(1, src.max_column + 1):
            assert backup.cell(row=r, column=c).value == src.cell(row=r, column=c).value

    assert counts["AP"] == 2  # includes the "(1)" duplicate — sort doesn't dedupe
    assert counts["Graphics"] == 1
    assert counts["Getty Unknown"] == 1
    assert counts["Getty Videos"] == 1
    assert counts["Shutterstock"] == 1
    assert counts["3rd parties"] == 1


def test_sort_workbook_custom_name_column(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path), name_column="Clip Name")
    assert sum(counts.values()) == 7
