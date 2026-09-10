"""Tests for the Getty Videos/Getty Stills repair command.

The fixtures reproduce the real shape of the problem: an old hand-checked
master that names its stills sheet "Getty pics" and uppercases bare ids,
and a new sorted workbook where those extension-less stills landed on
"Getty Videos".
"""
import pytest
from openpyxl import Workbook, load_workbook

from bigmedia.getty_split import fix_getty_split, pair_by_episode, episode_number

OLD_HEADERS = ["Name", "Source", "Clip Duration"]
NEW_HEADERS = ["Episode", "Clip Name", "Sequence In", "Clip Duration", "Source Reel Name"]


def _old_file(path, videos, pics, episode_in_name=True):
    wb = Workbook()
    wb.remove(wb.active)
    for title, names in [("Getty videos", videos), ("Getty pics", pics)]:
        ws = wb.create_sheet(title=title)
        for c, h in enumerate(OLD_HEADERS, start=1):
            ws.cell(row=1, column=c, value=h)
        for r, name in enumerate(names, start=2):
            ws.cell(row=r, column=1, value=name)
            ws.cell(row=r, column=3, value="00:00:02:00")
    wb.save(path)
    return path


def _new_file(path, episode, videos, stills=None, extra_sheets=()):
    wb = Workbook()
    wb.remove(wb.active)
    sheets = [("Getty Videos", videos)]
    if stills is not None:
        sheets.append(("Getty Stills", stills))
    for title, rows in list(extra_sheets) + sheets:
        ws = wb.create_sheet(title=title)
        for c, h in enumerate(NEW_HEADERS, start=1):
            ws.cell(row=1, column=c, value=h)
        for r, name in enumerate(rows, start=2):
            ws.cell(row=r, column=1, value=f"Episode {episode}")
            ws.cell(row=r, column=2, value=name)
            ws.cell(row=r, column=4, value="00:00:02:00")
            ws.cell(row=r, column=5, value="Getty Images - Footage")
    wb.save(path)
    return path


@pytest.fixture
def old(tmp_path):
    return _old_file(tmp_path / "EP4 - Submarine (kopie 2).xlsx",
                     videos=["GETTYIMAGES-111", "GETTYIMAGES-222"],
                     pics=["GETTYIMAGES-333", "GETTYIMAGES-444-.NEW.01"])


@pytest.fixture
def getty_split_pair(old, new):
    """Returns (old_path, new_path) tuple for styling tests."""
    return (old, new)


@pytest.fixture
def new(tmp_path):
    return _new_file(tmp_path / "SECRETS OF - SUBMARINE_sorted.xlsx", 4,
                     videos=["GettyImages-111.mov", "GETTYIMAGES-333", "GETTYIMAGES-333",
                             "GETTYIMAGES-444-.NEW.01", "GettyImages-999.mov"],
                     stills=["GettyImages-222.jpg"],
                     extra_sheets=[("Worksheet", ["anything.mov"]), ("AP", ["ap.mxf"])])


@pytest.fixture
def fixed(old, new, tmp_path):
    out = tmp_path / "fixed.xlsx"
    result = fix_getty_split(str(old), str(new), str(out))
    return load_workbook(out), result


def _names(ws):
    return [ws.cell(row=r, column=2).value for r in range(2, ws.max_row + 1)]


def test_stills_filed_as_videos_move_to_stills(fixed):
    wb, result = fixed
    assert result["video_to_stills"] == ["GETTYIMAGES-333", "GETTYIMAGES-444-.NEW.01"]
    # "-222.jpg" leaves for the videos sheet (the old file has it there),
    # so the stills sheet is exactly the three rows moved in from videos.
    assert _names(wb["Getty Stills"]) == [
        "GETTYIMAGES-333", "GETTYIMAGES-333", "GETTYIMAGES-444-.NEW.01"]


def test_every_use_of_a_misfiled_clip_moves_not_just_the_first(fixed):
    _, result = fixed
    assert result["rows_moved_to_stills"] == 3  # "-333" is used twice


def test_video_filed_as_still_moves_back_to_videos(fixed):
    wb, result = fixed
    # Old file has "-222" under Getty videos, the new one put it on stills.
    assert result["stills_to_video"] == ["GettyImages-222.jpg"]
    assert "GettyImages-222.jpg" in _names(wb["Getty Videos"])


def test_clip_the_old_version_does_not_have_is_left_alone(fixed):
    wb, result = fixed
    assert result["unmatched_videos"] == ["GettyImages-999.mov"]
    assert "GettyImages-999.mov" in _names(wb["Getty Videos"])


def test_matching_ignores_case_and_extension(fixed):
    wb, _ = fixed
    # "GETTYIMAGES-111" (old, videos) vs "GettyImages-111.mov" (new): same
    # clip, correctly filed, so it stays put.
    assert "GettyImages-111.mov" in _names(wb["Getty Videos"])


def test_other_sheets_are_untouched(fixed, new):
    wb, _ = fixed
    wb_src = load_workbook(new)
    assert wb.sheetnames == wb_src.sheetnames
    for title in ("Worksheet", "AP"):
        assert _names(wb[title]) == _names(wb_src[title])


def test_row_values_are_not_rewritten_only_the_sheet_changes(fixed):
    wb, _ = fixed
    ws = wb["Getty Stills"]
    row = [ws.cell(row=3, column=c).value for c in range(1, 6)]
    # "Source Reel Name" still says Footage -- that is EDL data, not ours.
    assert row == ["Episode 4", "GETTYIMAGES-333", None, "00:00:02:00", "Getty Images - Footage"]


def test_stills_sheet_is_created_when_the_new_file_has_none(old, tmp_path):
    new = _new_file(tmp_path / "no_stills_sorted.xlsx", 4,
                    videos=["GettyImages-111.mov", "GETTYIMAGES-333"])
    out = tmp_path / "created.xlsx"
    result = fix_getty_split(str(old), str(new), str(out))
    wb = load_workbook(out)
    assert result["stills_sheet_created"] is True
    assert wb.sheetnames == ["Getty Videos", "Getty Stills"]
    assert _names(wb["Getty Stills"]) == ["GETTYIMAGES-333"]
    assert _names(wb["Getty Videos"]) == ["GettyImages-111.mov"]


def test_no_getty_sheets_at_all_is_an_error(old, tmp_path):
    wb = Workbook()
    wb.active.title = "AP"
    wb.active.cell(row=1, column=2, value="Clip Name")
    path = tmp_path / "no_getty.xlsx"
    wb.save(path)
    with pytest.raises(ValueError):
        fix_getty_split(str(old), str(path), str(tmp_path / "out.xlsx"))


def test_grouped_totals_block_is_not_treated_as_a_clip(old, tmp_path):
    new = _new_file(tmp_path / "grouped_sorted.xlsx", 4, videos=["GETTYIMAGES-333"])
    wb = load_workbook(new)
    ws = wb["Getty Videos"]
    ws.cell(row=4, column=2, value="Total clips")
    ws.cell(row=5, column=2, value="Total Seconds")
    wb.save(new)
    result = fix_getty_split(str(old), str(new), str(tmp_path / "out.xlsx"))
    assert result["unmatched_videos"] == []


def test_episode_number_from_the_episode_column(new):
    assert episode_number(new, load_workbook(new)) == 4


def test_episode_number_falls_back_to_the_filename(old):
    assert episode_number(old, load_workbook(old)) == 4


def test_pairing_is_by_episode_number_not_filename(old, new, tmp_path):
    other_old = _old_file(tmp_path / "EP20 - Pyramids (kopie).xlsx", videos=["GETTYIMAGES-555"], pics=[])
    pairs, unpaired_old, unpaired_new = pair_by_episode([old, other_old], [new])
    assert pairs == [(4, old, new)]
    assert unpaired_old == [other_old]
    assert unpaired_new == []


def test_getty_split_output_is_styled_zebra_kept(getty_split_pair, tmp_path):
    """Getty output sheets get standard styling: black/white header,
    uniform font, fit-to-width. Zebra row fills are preserved."""
    old_p, new_p = getty_split_pair
    out = tmp_path / "fixed.xlsx"
    fix_getty_split(str(old_p), str(new_p), str(out))
    wb = load_workbook(str(out))
    getty_sheets = [t for t in wb.sheetnames if t.lower().startswith("getty")]
    assert getty_sheets, "No Getty sheets found"
    for t in getty_sheets:
        ws = wb[t]
        # Header cell (row 1, col 1) has black fill and white bold font
        h = ws.cell(row=1, column=1)
        assert h.fill.fgColor.rgb == "FF000000", f"{t}: header fill not black"
        assert h.font.bold, f"{t}: header not bold"
        assert h.font.color.rgb == "FFFFFFFF", f"{t}: header text not white"
        # Column A width is within expected range
        assert 10 <= ws.column_dimensions["A"].width <= 60, f"{t}: column A width {ws.column_dimensions['A'].width} out of range"
        # Zebra striping preserved: rows 2 and 3 have different fills (when present)
        if ws.max_row >= 3:
            assert ws.cell(row=2, column=1).fill != ws.cell(row=3, column=1).fill, \
                f"{t}: zebra striping lost (rows 2 and 3 have same fill)"
