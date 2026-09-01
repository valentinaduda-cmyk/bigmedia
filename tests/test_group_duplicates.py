from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
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
    # Total Seconds / Total Seconds >4s are live formulas over the Seconds
    # column (rows 2-10), not hardcoded numbers, so edits to the sheet
    # recompute them.
    seconds_col_letter = get_column_letter(seconds_col)
    assert ws.cell(row=15, column=name_col).value == "Total Seconds"
    assert ws.cell(row=15, column=name_col + 1).value == f"=SUM({seconds_col_letter}2:{seconds_col_letter}10)"
    assert ws.cell(row=16, column=name_col).value == "Total Seconds for clips >4s"
    assert ws.cell(row=16, column=name_col + 1).value == f'=SUMIF({seconds_col_letter}2:{seconds_col_letter}10,">4")'

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


def test_group_duplicates_matches_getty_pics_as_getty_stills_alias(tmp_path):
    # Real delivery files sometimes name the stills sheet "Getty pics"
    # instead of "Getty Stills" -- must be picked up by the default
    # sheet list without requiring --sheets.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Getty pics"
    ws.append(["Clip Name", "Clip Duration"])
    ws.append(["GettyImages-1.jpg", "00:00:00:10"])
    ws.append(["GettyImages-1 (1).jpg", "00:00:00:10"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path))

    assert set(results.keys()) == {"Getty pics"}
    assert results["Getty pics"]["total_clips"] == 2
    assert results["Getty pics"]["total_unique_clips"] == 1


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


def test_group_duplicates_falls_back_to_duration_column(tmp_path):
    # Some real delivery files title the duration column "Duration" instead
    # of "Clip Duration" (the default). group_duplicates_workbook should
    # still find it without requiring duration_column to be passed explicitly.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "AP"
    ws.append(["Clip Name", "Duration"])
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


def test_group_duplicates_collapses_codec_suffix_variant(tmp_path):
    # "GettyImages-1479649397.mov" and "GettyImages-1479649397_Apple_ProRes_422.mov"
    # are the same clip -- a codec tag baked in on re-export, not a different id.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Getty Videos"
    ws.append(["Clip Name", "Clip Duration"])
    ws.append(["GettyImages-1479649397.mov", "00:00:00:10"])
    ws.append(["GettyImages-1479649397_Apple_ProRes_422.mov", "00:00:00:10"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path), sheets=["Getty Videos"])

    assert results["Getty Videos"]["total_clips"] == 2
    assert results["Getty Videos"]["total_unique_clips"] == 1


def test_group_duplicates_collapses_getty_upscale_suffix_variant(tmp_path):
    # Same Getty clip, exported with an optional "_S000" segment marker plus
    # an "_upscaleNN" tag -- not different clips.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Getty Videos"
    ws.append(["Clip Name", "Clip Duration"])
    ws.append(["GettyImages-650878972_S000_upscale01.mov", "00:00:00:10"])
    ws.append(["GettyImages-650878972_S001_upscale02.mov", "00:00:00:10"])
    ws.append(["GettyImages-650878972_S002_upscale03.mov", "00:00:00:10"])
    ws.append(["GettyImages-650878972_upscale01.mov", "00:00:00:10"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path), sheets=["Getty Videos"])

    assert results["Getty Videos"]["total_clips"] == 4
    assert results["Getty Videos"]["total_unique_clips"] == 1


def test_group_duplicates_collapses_getty_render_prob4_comp_variant(tmp_path):
    # Getty QC/review-export junk ("Render", "Render <N>", "_prob4",
    # "_comp") re-appends the extension on every pass -- these are still
    # the same source clip and must group together, not split apart.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Getty Videos"
    ws.append(["Clip Name", "Clip Duration"])
    ws.append(["GettyImages-1142748666.mov Render 1_prob4.mov", "00:00:00:10"])
    ws.append(["GettyImages-1142748666.mov Render_prob4.mov", "00:00:00:10"])
    ws.append(["GettyImages-2064403302_comp", "00:00:00:10"])
    ws.append(["GettyImages-2064403302_comp_2", "00:00:00:10"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path), sheets=["Getty Videos"])

    assert results["Getty Videos"]["total_clips"] == 4
    assert results["Getty Videos"]["total_unique_clips"] == 2


def test_group_duplicates_reuters_groups_by_tape_id(tmp_path):
    # Reuters rushes clips: same tape ("m<digits>" embedded in the name) but
    # different timecode ranges cut from it -- still the same source clip
    # for grouping purposes, unlike the default full-name dedup key.
    src_path = tmp_path / "src.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Reuters"
    ws.append(["Clip Name", "Clip Duration"])
    names = [
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE ONE - 9_11 compilation_m909021_0-20-15-0-29-40.mov",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE ONE - 9_11 compilation_m909021_0-36-15-0-37-40.mov",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE ONE - 9_11 compilation_m909021_0-14-39-0-15-37 (1).mp4",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE THREE - 9_11 compilation_m917698_0-00-00-0-01-15_upscale.mp4",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE THREE - 9_11 compilation_m917698_0-03-15-0-03-35.mov",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE THREE - 9_11 compilation_m917698_0-08-00-0-08-18.mov",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE THREE - 9_11 compilation_m917698_0-32-45-0-33-05.mov",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE TWO - 9_11 compilation_m1241025_0-01-15-0-01-35.mov",
        "Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE TWO - 9_11 compilation_m1241025_0-35-40-0-36-00.mov",
    ]
    for name in names:
        ws.append([name, "00:00:00:10"])
    wb.save(src_path)

    out_path = tmp_path / "out.xlsx"
    results = group_duplicates_workbook(str(src_path), str(out_path), sheets=["Reuters"])

    assert results["Reuters"]["total_clips"] == 9
    assert results["Reuters"]["total_unique_clips"] == 3  # tapes m909021, m917698, m1241025


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


def _widths_by_header(ws):
    out = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h is None:
            continue
        d = ws.column_dimensions.get(get_column_letter(c))
        out[h] = d.width if d else None
    return out


def test_group_autofit_gives_each_header_the_same_width_on_every_sheet(sample_sorted_master, tmp_path):
    # Grouping inserts "Total Duration"/"Seconds" mid-table, so the same
    # header sits at a different column index on processed vs pass-through
    # sheets. Widths are therefore keyed by header name -- "Clip Name" must
    # be equally readable whichever tab you open.
    out_path = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out_path), autofit=True)

    wb = load_workbook(out_path)
    seen = {}
    for title in wb.sheetnames:
        if title == "Worksheet":
            continue
        for header, width in _widths_by_header(wb[title]).items():
            assert width is not None, f"{title}/{header} got no width"
            if header in seen:
                assert seen[header] == width, f"{header} differs on {title}"
            else:
                seen[header] = width
    assert "Clip Name" in seen


def test_group_autofit_respects_min_and_max(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out_path),
                              autofit=True, min_width=12, max_width=25)

    wb = load_workbook(out_path)
    for header, width in _widths_by_header(wb["Getty Videos"]).items():
        assert 12 <= width <= 25, f"{header} = {width}"


def test_group_uniform_font_on_data_cells(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out_path), uniform_font=True)

    wb = load_workbook(out_path)
    seen = set()
    for title in wb.sheetnames:
        if title == "Worksheet":
            continue
        ws = wb[title]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                f = ws.cell(row=r, column=c).font
                seen.add((f.name, f.size, f.italic))
    assert len(seen) == 1, f"data cells use {len(seen)} fonts: {seen}"


def test_group_formatting_leaves_worksheet_backup_alone(sample_sorted_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out_path),
                              autofit=True, uniform_font=True)

    src = load_workbook(sample_sorted_master)["Worksheet"]
    backup = load_workbook(out_path)["Worksheet"]
    assert backup.column_dimensions["B"].width == src.column_dimensions["B"].width


def test_group_without_formatting_flags_is_unchanged(sample_sorted_master, tmp_path):
    # Regression guard: the formatting options are opt-in only.
    a = tmp_path / "a.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(a))
    wb = load_workbook(a)
    ws = wb["Getty Videos"]
    assert ws.column_dimensions["B"].width == 45


def test_group_uniform_header_gives_every_header_cell_the_same_fill(tmp_path):
    # Real delivery headers are white bold text with a fill on only some
    # columns -- the unfilled ones render as invisible white-on-white, and
    # the inserted Total Duration/Seconds columns end up looking like the
    # only headers on the sheet.
    headers = ["Track", "Clip Name", "Clip Duration", "Source Reel Name"]
    src = tmp_path / "src.xlsx"
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Getty Videos")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFFFF")
        # Only the duration column carries a fill, like the real files.
        if h == "Clip Duration":
            cell.fill = PatternFill("solid", fgColor="FF000000")
    ws.cell(row=2, column=1, value="V1")
    ws.cell(row=2, column=2, value="GettyImages-1.mov")
    ws.cell(row=2, column=3, value="00:00:05:00")
    ws2 = wb.create_sheet("GFX")
    for c, h in enumerate(headers, start=1):
        ws2.cell(row=1, column=c, value=h)
    ws2.cell(row=2, column=2, value="7045_graphic.mov")
    wb.save(src)

    out = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(src), str(out), uniform_header=True)

    wbo = load_workbook(out)
    fills = set()
    for title in wbo.sheetnames:
        if title == "Worksheet":
            continue
        ws = wbo[title]
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=1, column=c).value is None:
                continue
            f = ws.cell(row=1, column=c).fill
            fills.add((f.patternType, f.fgColor.rgb))
    assert len(fills) == 1, f"header cells use {len(fills)} fills: {fills}"
    assert fills == {("solid", "FF000000")}


def test_group_uniform_header_is_opt_in(sample_sorted_master, tmp_path):
    out = tmp_path / "out.xlsx"
    group_duplicates_workbook(str(sample_sorted_master), str(out))
    ws = load_workbook(out)["Getty Videos"]
    # Unchanged default behaviour: header fill is whatever came through.
    assert ws.cell(row=1, column=1).fill.patternType is None
