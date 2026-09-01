import pytest
from openpyxl import Workbook, load_workbook
from bigmedia.sort_workbook import sort_workbook
from bigmedia.xlsx_utils import find_column


def test_sort_workbook_creates_backup_and_category_sheets(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path))

    wb = load_workbook(out_path)

    # Worksheet backup must be first and match the source exactly, including
    # "Source Reel Name" — sort no longer touches that column.
    assert wb.sheetnames[0] == "Worksheet"
    src = load_workbook(sample_master).active
    backup = wb["Worksheet"]
    for r in range(1, src.max_row + 1):
        for c in range(1, src.max_column + 1):
            assert backup.cell(row=r, column=c).value == src.cell(row=r, column=c).value

    assert counts["AP"] == 2  # includes the "(1)" duplicate — sort doesn't dedupe
    assert counts["GFX"] == 1
    assert counts["Getty Unknown"] == 1
    assert counts["Getty Videos"] == 1
    assert counts["Shutterstock"] == 1
    assert counts["3rd parties"] == 1


def test_sort_workbook_leaves_source_reel_name_untouched(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out_path))

    wb = load_workbook(out_path)
    src = load_workbook(sample_master).active
    source_col_idx = find_column(src, "Source Reel Name")
    name_col_idx = find_column(src, "Clip Name")

    for cat in ["AP", "GFX", "Getty Unknown", "Getty Videos", "Shutterstock", "3rd parties"]:
        ws = wb[cat]
        for r in range(2, ws.max_row + 1):
            name = ws.cell(row=r, column=name_col_idx).value
            for src_r in range(2, src.max_row + 1):
                if src.cell(row=src_r, column=name_col_idx).value == name:
                    assert ws.cell(row=r, column=source_col_idx).value == src.cell(row=src_r, column=source_col_idx).value
                    break


def test_sort_workbook_custom_name_column(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path), name_column="Clip Name")
    assert sum(counts.values()) == 7


def test_sort_workbook_falls_back_to_name_column(tmp_path):
    # Some real delivery files title the clip-name column "Name" instead of
    # "Clip Name" (the default). sort_workbook should still find it without
    # requiring name_column to be passed explicitly.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Source Reel Name"])
    ws.append(["apus_clip1.mxf", ""])
    ws.append(["shutterstock_777.mp4", ""])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(src_path), str(out_path))
    assert counts["AP"] == 1
    assert counts["Shutterstock"] == 1


def test_sort_workbook_skip_categories_omits_sheet_and_falls_back(sample_master, tmp_path):
    # A category turned off keeps its classification rules — its clips just
    # land on "3rd parties" for manual review instead of getting their own
    # sheet. Nothing is dropped.
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path), skip_categories=["Shutterstock"])

    wb = load_workbook(out_path)
    assert "Shutterstock" not in wb.sheetnames
    assert "3rd parties" in counts
    assert "Shutterstock" not in counts
    assert counts["3rd parties"] == 2  # its own row + the shutterstock one
    assert counts["AP"] == 2
    assert sum(counts.values()) == 7  # every non-empty row still accounted for


def test_sort_workbook_categories_allowlist(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path), categories=["AP", "GFX"])

    wb = load_workbook(out_path)
    assert wb.sheetnames == ["Worksheet", "AP", "GFX", "3rd parties"]
    assert counts["AP"] == 2
    assert counts["GFX"] == 1
    # Getty Unknown + Getty Videos + Shutterstock + the already-unclassified row
    assert counts["3rd parties"] == 4


def test_sort_workbook_category_names_are_case_insensitive(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    counts = sort_workbook(str(sample_master), str(out_path), skip_categories=["shutterstock", "GETTY VIDEOS"])
    assert "Shutterstock" not in counts
    assert "Getty Videos" not in counts
    assert counts["3rd parties"] == 3


def test_sort_workbook_rejects_both_selection_flags(sample_master, tmp_path):
    with pytest.raises(ValueError):
        sort_workbook(str(sample_master), str(tmp_path / "out.xlsx"),
                      categories=["AP"], skip_categories=["GFX"])


def test_sort_workbook_rejects_unknown_category_name(sample_master, tmp_path):
    # A typo must fail loudly rather than silently sorting nothing out.
    with pytest.raises(ValueError) as exc:
        sort_workbook(str(sample_master), str(tmp_path / "out.xlsx"), skip_categories=["Fox News"])
    assert "Fox News" in str(exc.value)


def test_sort_workbook_cannot_skip_the_fallback_category(sample_master, tmp_path):
    # "3rd parties" is where every turned-off category's clips land, so it
    # can't itself be turned off.
    with pytest.raises(ValueError):
        sort_workbook(str(sample_master), str(tmp_path / "out.xlsx"), skip_categories=["3rd parties"])


def test_sort_workbook_autofit_widths_are_uniform_across_sheets(sample_master, tmp_path):
    # Delivery files arrive with columns too narrow to read; autofit sizes
    # each column to its longest value so nobody has to drag column edges by
    # hand. Widths must be identical on every category sheet, otherwise the
    # workbook looks different depending on which tab you're on.
    out_path = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out_path), autofit=True)

    wb = load_workbook(out_path)
    sheets = [s for s in wb.sheetnames if s != "Worksheet"]
    ref = {k: v.width for k, v in wb[sheets[0]].column_dimensions.items()}
    assert ref, "autofit set no widths"
    for s in sheets[1:]:
        got = {k: v.width for k, v in wb[s].column_dimensions.items()}
        assert got == ref, f"{s} has different widths"


def test_sort_workbook_autofit_respects_min_and_max(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out_path), autofit=True,
                  min_width=12, max_width=30)

    wb = load_workbook(out_path)
    widths = [d.width for d in wb["AP"].column_dimensions.values()]
    assert widths
    for w in widths:
        assert 12 <= w <= 30

    # Column A holds the long clip names and must be capped, not unbounded.
    assert wb["AP"].column_dimensions["A"].width == 30


def test_sort_workbook_autofit_fits_the_longest_value_in_the_column(tmp_path):
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Clip Name", "Notes"])
    ws.append(["apus_clip1.mxf", "x"])
    ws.append(["shutterstock_777.mp4", "a much longer note than the header"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    sort_workbook(str(src_path), str(out_path), autofit=True, min_width=5, max_width=100)

    wb = load_workbook(out_path)
    b = wb["Shutterstock"].column_dimensions["B"].width
    assert b >= len("a much longer note than the header")


def test_sort_workbook_uniform_font_on_data_cells(sample_master, tmp_path):
    # Source files mix fonts column to column; the output should read as one
    # table. Header styling is left alone -- only data cells are normalized.
    out_path = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out_path), uniform_font=True)

    wb = load_workbook(out_path)
    seen = set()
    for s in wb.sheetnames:
        if s == "Worksheet":
            continue
        ws = wb[s]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                f = ws.cell(row=r, column=c).font
                seen.add((f.name, f.size, f.bold, f.italic))
    assert len(seen) == 1, f"data cells use {len(seen)} different fonts: {seen}"


def test_sort_workbook_worksheet_backup_is_untouched_by_formatting(sample_master, tmp_path):
    # The backup tab is a verbatim copy of the input -- autofit and font
    # normalization must not reach into it.
    out_path = tmp_path / "out.xlsx"
    sort_workbook(str(sample_master), str(out_path), autofit=True, uniform_font=True)

    src = load_workbook(sample_master).active
    backup = load_workbook(out_path)["Worksheet"]
    for col_letter in ("A", "B", "C", "D"):
        assert backup.column_dimensions[col_letter].width == src.column_dimensions[col_letter].width
