from bigmedia.dedupe import dedup_key, dedupe_workbook


def test_dedup_key_strips_extension_and_copy_marker():
    assert dedup_key("Clip_A.mov") == "Clip_A"
    assert dedup_key("Clip_A (1).mxf") == "Clip_A"
    assert dedup_key("Clip_A") == "Clip_A"


def test_dedupe_workbook_splits_kept_and_dropped(sample_master, tmp_path):
    out_path = tmp_path / "out.xlsx"
    result = dedupe_workbook(str(sample_master), str(out_path))
    # fixture has 7 rows, one of which is a "(1)" duplicate of another
    assert result["kept"] == 6
    assert result["dropped"] == 1
