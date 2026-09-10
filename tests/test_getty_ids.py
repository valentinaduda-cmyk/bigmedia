"""
Every real Getty clip name we've confirmed an id-extraction result for
lives here — the regression net for extract_getty_id(). Add new cases here
first if a new naming shape shows up, then adjust extract_getty_id() to
match, then rerun the whole file.
"""
import pytest
from openpyxl import Workbook, load_workbook

from bigmedia.getty_ids import extract_getty_id, clean_episode_title, build_getty_id_report

CASES = [
    ("GETTYIMAGES-1092-77", "1092-77"),
    ("GETTYIMAGES-1091006198", "1091006198"),
    ("GETTYIMAGES-96410456-", "96410456"),
    ("GETTYIMAGES-116718093_1", "116718093_1"),
    ("GETTYIMAGES-1B011400_0012 2", "1B011400_0012 2"),
    ("GETTYIMAGES-805-74 2", "805-74 2"),
    ("GETTYIMAGES-91953597.MOV", "91953597"),
    ("GettyImages-1001500162.mov", "1001500162"),
    ("GettyImages-2205665616_Apple_ProRes_422.mov", "2205665616"),
    ("GettyImages-2205665616_Apple_ProRes_422.jpg", "2205665616"),
    ("GettyImages-mr_00108323.mov", "mr_00108323"),
    ("GettyImages-1344-77.mov", "1344-77"),
    ("GettyImages-1B02673_0003.mov", "1B02673_0003"),
    ("GettyImages-mr_00097951.mov", "mr_00097951"),
    ("GETTYIMAGES-MR_00061694", "MR_00061694"),
    ("GETTYIMAGES-MR_00061694.mov", "MR_00061694"),
    ("GETTYIMAGES-MR_00061694.mp4", "MR_00061694"),
    ("GETTYIMAGES-MR_00061694.jpg", "MR_00061694"),
    ("GettyImages-1290166256.mov 25", "1290166256"),
    ("GettyImages-1328879201.mp4 25", "1328879201"),
    ("GettyImages-1053081380-NTSC.mov", "1053081380-NTSC"),
    ("GettyImages-468-9-PAL.mov", "468-9-PAL"),
    ("GettyImages-815292152_Denoise_02.mov", "815292152"),
    ("GettyImages-1150740782.new.02", "1150740782"),
    ("GettyImages-1297644901.mov.new.01", "1297644901"),
    ("GettyImages-1339022866.new.03", "1339022866"),
    ("GettyImages-518262576.new.05", "518262576"),
    # Real ids can end in a letter+digits tag that looks superficially like
    # a junk suffix (e.g. "_Denoise") but is actually part of the id itself
    # -- must survive extension-cutting, not just the no-extension path.
    ("GettyImages-1B010728_t010", "1B010728_t010"),
    ("GettyImages-1B010728_t010.mov", "1B010728_t010"),
    ("GettyImages-1B010728_t008.mov", "1B010728_t008"),
    ("GettyImages-1B03368_0006.mov", "1B03368_0006"),
    # Known junk tags must still be stripped, including when preceded or
    # followed by a legitimate-looking segment.
    ("GettyImages-650878972_S000_upscale01.mov", "650878972"),
    ("GettyImages-650878972_upscale01.mov", "650878972"),
    # Newer export/post-process tags: "_SM"/"_SMnn", "_AIUPSCALE",
    # "_APPLEPRORESHQ" (no separators), "_DFR"/"_DFRnn"/"_DFR_OK", and the
    # "-640_ADPP" resolution block (optionally followed by a "(N)" copy
    # marker before the extension). Stripped with or without an extension,
    # and in combination.
    ("GettyImages-1234567890_SM.mov", "1234567890"),
    ("GettyImages-1234567890_SM01.mov", "1234567890"),
    ("GettyImages-1234567890_SM02", "1234567890"),
    ("GettyImages-1234567890_AIUPSCALE.mov", "1234567890"),
    ("GettyImages-1234567890_APPLEPRORESHQ", "1234567890"),
    ("GettyImages-1234567890_APPLEPRORESHQ.MOV", "1234567890"),
    ("GettyImages-1234567890_DFR", "1234567890"),
    ("GettyImages-1234567890_DFR01", "1234567890"),
    ("GettyImages-1234567890_DFR_OK", "1234567890"),
    ("GettyImages-1234567890_DFR01_OK", "1234567890"),
    ("GettyImages-1234567890_DFR02_OK.mov", "1234567890"),
    ("GETTYIMAGES-2217807802-640_ADPP.MP4", "2217807802"),
    ("GETTYIMAGES-2217807802-640_ADPP (3).MP4", "2217807802"),
    ("GETTYIMAGES-2217807802-640_ADPP(4).MP4", "2217807802"),
    ("GettyImages-1234567890_SM01_APPLEPRORESHQ.MOV", "1234567890"),
    ("GettyImages-1234567890_DFR01_OK_SM02.MP4", "1234567890"),
    ("GETTYIMAGES-2217807802-640_ADPP_AIUPSCALE (2).MP4", "2217807802"),
]


@pytest.mark.parametrize("name,expected", CASES, ids=[c[0] for c in CASES])
def test_extract_getty_id(name, expected):
    assert extract_getty_id(name) == expected


TITLE_CASES = [
    ("EP1 - Eiffel Tower (kopie).xlsx", "EP1 - Eiffel Tower"),
    ("EP2 - CDG (kopie).xlsx", "EP2 - CDG"),
    ("EP3 - Versailles.xlsx", "EP3 - Versailles"),
    ("EP1_master_sorted.xlsx", "EP1_master_sorted"),
]


@pytest.mark.parametrize("filename,expected", TITLE_CASES)
def test_clean_episode_title(filename, expected):
    assert clean_episode_title(filename) == expected


def _make_sorted_workbook(path, rows, stills_rows=None):
    """rows/stills_rows: list of (name, seconds) tuples. `seconds` mirrors
    the real convention: it's blank (None) on every row of a duplicate run
    except the last, which carries that clip's total-duration-in-seconds."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty Videos")
    ws.append(["Track", "Clip Name", "Seconds"])
    for name, seconds in rows:
        ws.append(["V1", name, seconds])
    if stills_rows is not None:
        ws_stills = wb.create_sheet(title="Getty Stills")
        ws_stills.append(["Track", "Clip Name", "Seconds"])
        for name, seconds in stills_rows:
            ws_stills.append(["V1", name, seconds])
    ws2 = wb.create_sheet(title="AP")  # not a Getty sheet, must be ignored
    ws2.append(["Track", "Clip Name", "Seconds"])
    ws2.append(["V1", "BM1234_something.mxf", 10])
    wb.save(path)


def test_build_getty_id_report_writes_header_fields(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(
        str(in_dir), str(out_path), project_name="Top 10 Secrets of Technology",
        production_company="KM Record a.s./Big Media", broadcaster="Discovery",
        rights="in perpetuity/worldwide/all media",
    )

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=1, column=1).value == "Customer Declaration Form"
    assert ws.cell(row=2, column=1).value == "Production Company:"
    assert ws.cell(row=2, column=2).value == "KM Record a.s./Big Media"
    assert ws.cell(row=3, column=1).value == "Project Name:"
    assert ws.cell(row=3, column=2).value == "Top 10 Secrets of Technology"
    assert ws.cell(row=4, column=1).value == "Broadcaster:"
    assert ws.cell(row=4, column=2).value == "Discovery"
    assert ws.cell(row=5, column=1).value == "Rights Requested:"
    assert ws.cell(row=5, column=2).value == "in perpetuity/worldwide/all media"


def test_build_getty_id_report_header_field_defaults(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=2, column=2).value == "KM Record a.s./Big Media"
    assert ws.cell(row=4, column=2).value is None  # openpyxl normalizes empty strings to None
    assert ws.cell(row=5, column=2).value == "in perpetuity/worldwide/all media"


def test_build_getty_id_report_consolidates_multiple_files(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1_master_sorted.xlsx", [
        ("GettyImages-1001500162.mov", 6),
        ("GettyImages-1344-77.mov", 5),
    ])
    _make_sorted_workbook(in_dir / "EP2_master_sorted.xlsx", [
        ("GettyImages-mr_00108323.mov", 7),
    ])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    assert counts == {
        "EP1_master_sorted.xlsx": {"video": 2, "stills": 0},
        "EP2_master_sorted.xlsx": {"video": 1, "stills": 0},
    }

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]

    # First block: title, subtitle, Asset ID/Duration headers, formula row, then data.
    assert ws.cell(row=7, column=1).value == "EP1_master_sorted"
    assert ws.cell(row=8, column=1).value == "Getty Images Video"
    assert ws.cell(row=9, column=1).value == "Asset ID"
    assert ws.cell(row=9, column=2).value == "Duration"
    assert ws.cell(row=10, column=1).value == "=COUNTA(A11:A12)"
    assert ws.cell(row=10, column=2).value == "=SUM(B11:B12)"
    assert ws.cell(row=11, column=1).value == "1001500162"
    assert ws.cell(row=11, column=2).value == 6
    assert ws.cell(row=12, column=1).value == "1344-77"
    assert ws.cell(row=12, column=2).value == 5


def test_build_getty_id_report_dedupes_and_filters_by_seconds(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [
        ("GettyImages-1000000001.mov", 3),          # unique, too short -> excluded
        ("GettyImages-2000000002.mov", None),        # dup run: blank until last row
        ("GettyImages-2000000002.mov", None),
        ("GettyImages-2000000002.mov", 6),           # aggregate seconds -> one entry
        ("GettyImages-3000000003.mov", 5),           # unique, exactly at threshold -> included
    ])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 2, "stills": 0}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=11, column=1).value == "2000000002"
    assert ws.cell(row=11, column=2).value == 6
    assert ws.cell(row=12, column=1).value == "3000000003"
    assert ws.cell(row=12, column=2).value == 5


def test_build_getty_id_report_skips_garbled_seconds_cells(tmp_path):
    # A real file had a corrupted, non-numeric value in a "Seconds" cell —
    # must be treated as "doesn't qualify", not crash the whole batch.
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [
        ("GettyImages-1000000001.mov", "#REF!"),
        ("GettyImages-2000000002.mov", 6),
    ])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}


def test_build_getty_id_report_stills_ignore_duration_threshold(tmp_path):
    # Stills have no meaningful "duration" -- min/max-seconds must only
    # filter the video sheet; every uniquely-named still is extracted
    # regardless of what's in its Seconds cell (including blank/None).
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[
            ("GettyImages-1000000001.mov", 3),  # below default min_seconds=5 -> excluded
            ("GettyImages-2000000002.mov", 6),
        ],
        stills_rows=[
            ("GettyImages-3000000003.jpg", 1),     # short duration, still included
            ("GettyImages-4000000004.jpg", None),  # blank duration, still included
        ],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 2}}


def test_build_getty_id_report_min_seconds_is_configurable(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1000000001.mov", 3)])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X", min_seconds=0)
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}


def test_build_getty_id_report_max_seconds_is_configurable(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [
        ("GettyImages-1000000001.mov", 3),   # below 5 -> included when max_seconds=5, min_seconds=0
        ("GettyImages-2000000002.mov", 6),   # at/above 5 -> excluded when max_seconds=5
    ])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(
        str(in_dir), str(out_path), project_name="X", min_seconds=0, max_seconds=5,
    )
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=11, column=1).value == "1000000001"


def test_build_getty_id_report_max_seconds_default_is_unbounded(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1000000001.mov", 9999)])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}


def test_build_getty_id_report_falls_back_to_name_column(tmp_path):
    # Real delivery files disagree: some use "Clip Name", others plain
    # "Name", even within the same batch. Neither should need -o/--name-column.
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", [("GettyImages-1001500162.mov", 6)])

    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append([" Status", "Location", "Track", "Name", "Source", "Seconds"])
    ws.append([None, None, None, "GettyImages-1344-77.mov", "Getty", 5])
    wb.save(in_dir / "EP2.xlsx")

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {
        "EP1.xlsx": {"video": 1, "stills": 0},
        "EP2.xlsx": {"video": 1, "stills": 0},
    }

    wb_out = load_workbook(out_path)
    ws_out = wb_out["Getty IDs"]
    assert ws_out.cell(row=11, column=1).value == "1001500162"
    assert ws_out.cell(row=11, column=8).value == "1344-77"


def test_build_getty_id_report_ignores_blank_names(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    path = in_dir / "EP1.xlsx"
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append(["Track", "Clip Name", "Seconds"])
    ws.append(["V1", "GettyImages-1001500162.mov", 6])
    ws.append(["V1", None, None])
    wb.save(path)

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}


def test_build_getty_id_report_reads_stills_sheet(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1001500162.mov", 6)],
        stills_rows=[("GettyImages-2001500162.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 1}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]

    # Video block: columns 1-2 (A-B).
    assert ws.cell(row=7, column=1).value == "EP1"
    assert ws.cell(row=8, column=1).value == "Getty Images Video"
    assert ws.cell(row=11, column=1).value == "1001500162"
    assert ws.cell(row=11, column=2).value == 6

    # One blank spacer column (3 / C) between video and stills.
    assert ws.cell(row=7, column=3).value is None

    # Stills block: columns 4-5 (D-E).
    assert ws.cell(row=7, column=4).value == "EP1"
    assert ws.cell(row=8, column=4).value == "Getty Images Stills"
    assert ws.cell(row=9, column=4).value == "Asset ID"
    assert ws.cell(row=9, column=5).value == "Duration"
    assert ws.cell(row=10, column=4).value == "=COUNTA(D11:D11)"
    assert ws.cell(row=10, column=5).value == "=SUM(E11:E11)"
    assert ws.cell(row=11, column=4).value == "2001500162"
    assert ws.cell(row=11, column=5).value == 5


def test_build_getty_id_report_missing_stills_sheet_writes_empty_block(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", rows=[("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=8, column=4).value == "Getty Images Stills"
    assert ws.cell(row=10, column=4).value == "=COUNTA(D11:D11)"
    assert ws.cell(row=11, column=4).value is None


def test_build_getty_id_report_two_episodes_video_and_stills_column_math(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1000000001.mov", 6)],
        stills_rows=[("GettyImages-2000000001.jpg", 5)],
    )
    _make_sorted_workbook(
        in_dir / "EP2.xlsx",
        rows=[("GettyImages-1000000002.mov", 6)],
        stills_rows=[("GettyImages-2000000002.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    # EP1: video cols 1-2, stills cols 4-5. EP2: video cols 8-9, stills cols 11-12.
    assert ws.cell(row=11, column=1).value == "1000000001"
    assert ws.cell(row=11, column=4).value == "2000000001"
    assert ws.cell(row=11, column=8).value == "1000000002"
    assert ws.cell(row=11, column=11).value == "2000000002"

    # Two blank columns (6, 7 / F, G) between EP1's stills block and EP2's video block.
    assert ws.cell(row=7, column=6).value is None
    assert ws.cell(row=7, column=7).value is None
    assert ws.cell(row=7, column=8).value == "EP2"


def test_build_getty_id_report_falls_back_to_getty_pics_stills_sheet(tmp_path):
    # Real delivery files sometimes name the stills sheet "Getty pics"
    # instead of "Getty Stills" -- must be found without -o/--stills-sheet.
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append(["Track", "Clip Name", "Seconds"])
    ws.append(["V1", "GettyImages-1001500162.mov", 6])
    ws_stills = wb.create_sheet(title="Getty pics")
    ws_stills.append(["Track", "Clip Name", "Seconds"])
    ws_stills.append(["V1", "GettyImages-2001500162.jpg", 5])
    wb.save(in_dir / "EP1.xlsx")

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X")
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 1}}


def test_build_getty_id_report_custom_stills_sheet_name(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Getty videos")
    ws.append(["Track", "Clip Name", "Seconds"])
    ws.append(["V1", "GettyImages-1001500162.mov", 6])
    ws_stills = wb.create_sheet(title="Getty Stills")  # non-default name
    ws_stills.append(["Track", "Clip Name", "Seconds"])
    ws_stills.append(["V1", "GettyImages-2001500162.jpg", 5])
    wb.save(in_dir / "EP1.xlsx")

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(
        str(in_dir), str(out_path), project_name="X", stills_sheet_name="Getty Stills",
    )
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 1}}


def test_build_getty_id_report_videos_only(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1000000001.mov", 6)],
        stills_rows=[("GettyImages-2000000002.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X", include_stills=False)
    assert counts == {"EP1.xlsx": {"video": 1, "stills": 0}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=8, column=1).value == "Getty Images Video"
    assert ws.cell(row=11, column=1).value == "1000000001"
    # No stills block written -> nothing at the column the stills block
    # would otherwise occupy.
    assert ws.cell(row=8, column=4).value is None


def test_build_getty_id_report_stills_only(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1000000001.mov", 6)],
        stills_rows=[("GettyImages-2000000002.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    counts = build_getty_id_report(str(in_dir), str(out_path), project_name="X", include_video=False)
    assert counts == {"EP1.xlsx": {"video": 0, "stills": 1}}

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    assert ws.cell(row=8, column=1).value == "Getty Images Stills"
    assert ws.cell(row=11, column=1).value == "2000000002"


_INSTRUCTIONS_TEXT_START = "Please divide content into appropriate asset type."
_NOTES_TEXT_START = "Upon receipt of the Proposed Usage Declaration form"


def test_build_getty_id_report_text_boxes_one_file(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", rows=[("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    # EP1 alone: video cols 1-2, stills cols 4-5, last_used_col=5.
    # box_start = 5 + 1 + 2 (BOX_GAP) = 8 (H).
    assert ws.cell(row=5, column=8).value == "Instructions:"
    assert ws.cell(row=7, column=8).value.startswith(_INSTRUCTIONS_TEXT_START)
    # notes_start = 8 + 6 (INSTR width) + 1 (inner gap) = 15 (O).
    assert ws.cell(row=5, column=15).value == "Important Notes on Licensing:"
    assert ws.cell(row=7, column=15).value.startswith(_NOTES_TEXT_START)


def test_build_getty_id_report_text_boxes_shift_with_two_files(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(
        in_dir / "EP1.xlsx",
        rows=[("GettyImages-1000000001.mov", 6)],
        stills_rows=[("GettyImages-2000000001.jpg", 5)],
    )
    _make_sorted_workbook(
        in_dir / "EP2.xlsx",
        rows=[("GettyImages-1000000002.mov", 6)],
        stills_rows=[("GettyImages-2000000002.jpg", 5)],
    )

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    # EP1+EP2: video 8-9, stills 11-12, last_used_col=12.
    # box_start = 12 + 1 + 2 = 15 (O).
    assert ws.cell(row=5, column=15).value == "Instructions:"
    # notes_start = 15 + 6 + 1 = 22 (V).
    assert ws.cell(row=5, column=22).value == "Important Notes on Licensing:"


def test_build_getty_id_report_text_box_merge_ranges(tmp_path):
    in_dir = tmp_path / "in"
    in_dir.mkdir()
    _make_sorted_workbook(in_dir / "EP1.xlsx", rows=[("GettyImages-1001500162.mov", 6)])

    out_path = tmp_path / "getty_ids.xlsx"
    build_getty_id_report(str(in_dir), str(out_path), project_name="X")

    wb = load_workbook(out_path)
    ws = wb["Getty IDs"]
    merged = {str(r) for r in ws.merged_cells.ranges}
    assert "H5:M6" in merged    # Instructions header, 6 cols wide, 2 rows tall
    assert "H7:M32" in merged   # Instructions body, 6 cols wide, 26 rows tall
    assert "O5:V6" in merged    # Notes header, 8 cols wide, 2 rows tall
    assert "O7:V32" in merged   # Notes body, 8 cols wide, 26 rows tall
