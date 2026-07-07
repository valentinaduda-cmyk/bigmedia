from openpyxl import load_workbook
from bigmedia.dedupe import dedup_key, dedupe_workbook
from bigmedia.xlsx_utils import find_column


def test_dedup_key_strips_extension_and_copy_marker():
    assert dedup_key("Clip_A.mov") == "Clip_A"
    assert dedup_key("Clip_A (1).mxf") == "Clip_A"
    assert dedup_key("Clip_A") == "Clip_A"


def test_dedup_key_strips_trailing_number_tag_after_extension():
    # A real duplicate pair: same clip, one with a stray " 25" tag appended
    # after the extension (e.g. a frame-rate/version annotation).
    assert dedup_key("GettyImages-627-122.mov 25") == dedup_key("GettyImages-627-122.mov")
    assert dedup_key("GettyImages-627-122.mov 25") == "GettyImages-627-122"
    # Without an extension, a trailing " <digits>" is NOT noise -- it can be
    # a meaningful part of the id (e.g. "GettyImages-805-74 2"), so it must
    # be left alone.
    assert dedup_key("GettyImages-805-74 2") == "GettyImages-805-74 2"


def test_dedupe_workbook_splits_kept_and_dropped(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    result = dedupe_workbook(str(sample_master), str(out_path))
    # fixture has 7 rows, one of which is a "(1)" duplicate of another
    assert result["kept"] == 6
    assert result["dropped"] == 1


def test_dedupe_workbook_fills_source_reel_name(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    dedupe_workbook(str(sample_master), str(out_path))

    wb = load_workbook(out_path)
    for title in ["Deduped", "Duplicates"]:
        ws = wb[title]
        source_col_idx = find_column(ws, "Source Reel Name")
        for r in range(2, ws.max_row + 1):
            assert ws.cell(row=r, column=source_col_idx).value == title

    # Backup rows are labeled with whichever sheet they ended up on.
    src = load_workbook(sample_master).active
    backup = wb["Worksheet"]
    source_col_idx = find_column(src, "Source Reel Name")
    duplicate_row = 8  # the "(1)" copy of row 2, per the fixture
    assert backup.cell(row=duplicate_row, column=source_col_idx).value == "Duplicates"
    assert backup.cell(row=2, column=source_col_idx).value == "Deduped"
