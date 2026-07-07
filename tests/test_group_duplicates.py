from openpyxl import Workbook, load_workbook
from bigmedia.group_duplicates import group_duplicates_workbook
from bigmedia.xlsx_utils import find_column


def test_group_duplicates_reorders_and_sums_getty_videos(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(sample_sorted_master), str(out_path))

    wb = load_workbook(out_path)
    ws = wb["Getty Videos"]

    name_col = find_column(ws, "Clip Name")
    dur_col = find_column(ws, "Clip Duration")
    total_col = dur_col + 1  # inserted right after Clip Duration, not at the end
    seconds_col = dur_col + 2

    assert ws.cell(row=1, column=total_col).value == "Total Duration"
    assert ws.cell(row=1, column=seconds_col).value == "Seconds"
    # Columns after Clip Duration got pushed right by two.
    assert ws.cell(row=1, column=seconds_col + 1).value == "Source Reel Name"

    # Rows are reordered into ascending-by-id groups, duplicates adjacent.
    names = [ws.cell(row=r, column=name_col).value for r in range(2, 11)]
    assert names == [
        "GettyImages-0001.mov",
        "GettyImages-1073703562.mov",
        "GettyImages-2158057203.mov",
        "GettyImages-2158057203 (1).mxf",
        "GettyImages-2158057203 (2).mov",
        "GettyImages-356-30.mov",
        "GettyImages-356-31.mov",
        "GettyImages-9999.mov",
        "GettyImages-9999 (1).mxf",
    ]

    # Total Duration/Seconds only appear on the last row of each group.
    assert ws.cell(row=2, column=total_col).value == "00:00:00:20"
    assert ws.cell(row=2, column=seconds_col).value == 1

    assert ws.cell(row=3, column=total_col).value == "00:00:03:06"
    assert ws.cell(row=3, column=seconds_col).value == 4

    assert ws.cell(row=4, column=total_col).value is None  # mid-group
    assert ws.cell(row=5, column=total_col).value is None  # mid-group
    assert ws.cell(row=6, column=total_col).value == "00:00:05:03"  # 44+42+42 frames
    assert ws.cell(row=6, column=seconds_col).value == 6

    assert ws.cell(row=10, column=total_col).value == "00:00:03:00"  # 50+25 frames
    assert ws.cell(row=10, column=seconds_col).value == 3

    # Thick bottom border on the last row of each group, full row width.
    assert ws.cell(row=3, column=name_col).border.bottom.style == "thick"
    assert ws.cell(row=4, column=name_col).border.bottom.style != "thick"
    assert ws.cell(row=6, column=seconds_col).border.bottom.style == "thick"

    # Every row of a group is highlighted the same color, across the full
    # row width: GREEN if the group's total is 4s or less, YELLOW if over.
    # Group "2158057203" (rows 4-6) totals 6s -> yellow, on every row.
    assert ws.cell(row=4, column=name_col).fill.fgColor.rgb == "00FFFF00"
    assert ws.cell(row=5, column=name_col).fill.fgColor.rgb == "00FFFF00"
    assert ws.cell(row=6, column=seconds_col).fill.fgColor.rgb == "00FFFF00"
    # "GettyImages-1073703562.mov" totals exactly 4s -> not over 4s -> green.
    assert ws.cell(row=3, column=name_col).value == "GettyImages-1073703562.mov"
    assert ws.cell(row=3, column=name_col).fill.fgColor.rgb == "0092D050"
    assert ws.cell(row=3, column=1).fill.fgColor.rgb == "0092D050"  # full row width

    # 9 data rows (rows 2-10) -> row 11 is a blank spacer, summary starts row 12.
    assert ws.cell(row=11, column=name_col).value is None
    assert ws.cell(row=12, column=name_col).value == "Total clips"
    assert ws.cell(row=12, column=name_col + 1).value == 9
    assert ws.cell(row=13, column=name_col).value == "Total unique clips"
    assert ws.cell(row=13, column=name_col + 1).value == 6
    assert ws.cell(row=14, column=name_col).value == "Total clips >4s"
    assert ws.cell(row=14, column=name_col + 1).value == 1
    assert ws.cell(row=15, column=name_col).value == "Total Seconds"
    assert ws.cell(row=15, column=name_col + 1).value == 16
    assert ws.cell(row=16, column=name_col).value == "Total Seconds for clips >4s"
    assert ws.cell(row=16, column=name_col + 1).value == 6

    assert results["Getty Videos"] == {
        "total_clips": 9,
        "total_unique_clips": 6,
        "total_clips_over_4s": 1,
        "total_seconds": 16,
        "total_seconds_over_4s": 6,
    }


def test_group_duplicates_leaves_other_sheets_untouched(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out_path))

    src = load_workbook(sample_sorted_master)
    out = load_workbook(out_path)

    for title in ["GFX", "Worksheet"]:
        src_ws, out_ws = src[title], out[title]
        assert out_ws.max_column == src_ws.max_column  # no Total Duration/Seconds added
        for r in range(1, src_ws.max_row + 1):
            for c in range(1, src_ws.max_column + 1):
                assert out_ws.cell(row=r, column=c).value == src_ws.cell(row=r, column=c).value


def test_group_duplicates_only_processes_default_sheets(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(sample_sorted_master), str(out_path))
    assert set(results.keys()) == {"Getty Videos", "AP"}


def test_group_duplicates_matches_sheet_names_case_insensitively(sample_sorted_master, tmp_path):
    # Real delivery files sometimes title this sheet "getty videos" (any
    # case) instead of our own "Getty Videos" convention.
    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(sample_sorted_master), str(out_path), sheets=["ap", "GETTY VIDEOS"])
    assert set(results.keys()) == {"Getty Videos", "AP"}


def test_group_duplicates_custom_sheets_override(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(sample_sorted_master), str(out_path), sheets=["GFX"])
    assert set(results.keys()) == {"GFX"}


def test_group_duplicates_falls_back_to_name_column(tmp_path):
    # Some real delivery files title the clip-name column "Name" instead of
    # "Clip Name" (the default). group_duplicates_workbook should still find
    # it without requiring name_column to be passed explicitly.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "AP"
    ws.append(["Name", "Clip Duration"])
    ws.append(["apus_clip1.mxf", "00:00:00:10"])
    ws.append(["apus_clip1 (1).mxf", "00:00:00:10"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path), sheets=["AP"])

    assert results["AP"]["total_clips"] == 2
    assert results["AP"]["total_unique_clips"] == 1


def test_group_duplicates_collapses_trailing_number_tag_variant(tmp_path):
    # "GettyImages-627-122.mov" and "GettyImages-627-122.mov 25" are the
    # same clip -- a stray " 25" tag after the extension, not a different id.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Getty Videos"
    ws.append(["Clip Name", "Clip Duration"])
    ws.append(["GettyImages-627-122.mov", "00:00:00:10"])
    ws.append(["GettyImages-627-122.mov 25", "00:00:00:10"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path), sheets=["Getty Videos"])

    assert results["Getty Videos"]["total_clips"] == 2
    assert results["Getty Videos"]["total_unique_clips"] == 1


def test_group_duplicates_blank_duration_counts_as_zero(tmp_path):
    # Real delivery files sometimes leave Clip Duration blank on a row.
    # It should count as 0 frames, not crash the whole batch.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "AP"
    ws.append(["Clip Name", "Clip Duration"])
    ws.append(["apus_clip1.mxf", "00:00:00:10"])
    ws.append(["apus_clip1 (1).mxf", None])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path), sheets=["AP"])

    assert results["AP"]["total_clips"] == 2
    assert results["AP"]["total_unique_clips"] == 1
    assert results["AP"]["total_seconds"] == 1  # 10 frames @25fps -> ceil to 1s
