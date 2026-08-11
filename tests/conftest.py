"""Builds a small synthetic master.xlsx fixture on demand so tests don't
depend on committing real client data to the repo."""
import pytest
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


@pytest.fixture
def sample_master(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["Clip Name", "Duration", "Notes", "Source Reel Name"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="CCCCCC")

    rows = [
        ("BM45214-3_Athens_Celebrates_Liberation_NO_SOUND_HD_50_XDCAM_50.mxf", "00:01:20:00", "AP archive", "TBD"),
        ("045_SO_EP18_01_3DExplainer_TectonicPlates_TXLS.mov", "00:00:45:00", "graphics element", "TBD"),
        ("GettyImages-51240904", "", "unknown media type", ""),
        ("GettyImages-99999999.mov", "00:02:00:00", "getty video", "TBD"),
        ("shutterstock_777.mp4", "00:00:10:00", "", "wrong value"),
        ("random_file_xyz.mov", "00:00:05:00", "unclassified", ""),
        ("BM45214-3_Athens_Celebrates_Liberation_NO_SOUND_HD_50_XDCAM_50 (1).mxf", "00:01:20:00", "duplicate of row 2", "TBD"),
    ]
    for r, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=val if val != "" else None)
        ws.cell(row=r, column=2).number_format = "hh:mm:ss"

    for col_letter, width in [("A", 60), ("B", 15), ("C", 25), ("D", 20)]:
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"

    path = tmp_path / "sample_master.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def sample_sorted_master(tmp_path):
    """A small stand-in for the output of `bigmedia sort`: a multi-sheet
    workbook with the real-world column layout (Track, Clip Name, Enabled,
    Sequence In/Out, Clip Duration, Source Reel Name, Source In/Out, Source
    Duration), scrambled rows to prove grouping actually reorders them, and
    an untouched sheet to prove non-target sheets pass through unchanged."""
    headers = [
        "Track", "Clip Name", "Enabled", "Sequence In", "Sequence Out",
        "Clip Duration", "Source Reel Name", "Source In", "Source Out", "Source Duration",
    ]

    def add_sheet(wb, title, rows):
        ws = wb.create_sheet(title=title)
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="CCCCCC")
        for r, (name, duration) in enumerate(rows, start=2):
            ws.cell(row=r, column=1, value="V1")
            ws.cell(row=r, column=2, value=name)
            ws.cell(row=r, column=3, value="x")
            ws.cell(row=r, column=6, value=duration)
            ws.cell(row=r, column=7, value=title)
            ws.cell(row=r, column=10, value=duration)
            # Zebra-stripe by ORIGINAL row parity, so a test can prove the
            # grouping step recomputes stripes from each row's new position
            # rather than carrying its old color along.
            fill = PatternFill("solid", fgColor="FFFF00" if r % 2 == 0 else "00FF00")
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = fill
        ws.column_dimensions["B"].width = 45
        ws.column_dimensions["F"].width = 15
        ws.freeze_panes = "A2"
        return ws

    wb = Workbook()
    wb.remove(wb.active)

    # Deliberately scrambled: the two "GettyImages-2158057203" duplicates
    # are not adjacent, to prove grouping actually reorders rows.
    getty_rows = [
        ("GettyImages-1073703562.mov", "00:00:03:06"),   # single clip, >=4s
        ("GettyImages-2158057203.mov", "00:00:01:19"),   # dup group (3), >=4s combined
        ("GettyImages-356-31.mov", "00:00:00:10"),        # distinct id despite shared prefix
        ("GettyImages-9999.mov", "00:00:02:00"),          # dup group (2) via copy marker + ext change
        ("GettyImages-2158057203 (1).mxf", "00:00:01:17"),
        ("GettyImages-356-30.mov", "00:00:00:10"),        # distinct id despite shared prefix
        ("GettyImages-0001.mov", "00:00:00:20"),          # single clip, <4s
        ("GettyImages-9999 (1).mxf", "00:00:01:00"),
        ("GettyImages-2158057203 (2).mov", "00:00:01:17"),
    ]
    add_sheet(wb, "Getty Videos", getty_rows)
    add_sheet(wb, "AP", [("BM1234_something.mxf", "00:00:02:00")])
    add_sheet(wb, "GFX", [("7045_graphic.mov", "00:00:01:00")])
    add_sheet(wb, "Worksheet", getty_rows + [("BM1234_something.mxf", "00:00:02:00")])

    path = tmp_path / "sample_sorted_master.xlsx"
    wb.save(path)
    return path


# (clip name, sequence in, sequence out, clip duration, source duration)
THIRD_PARTIES_ROWS = [
    # Same clip used 3 times, rows scattered, earliest use is NOT the first row.
    ("Evan Fairbanks WTC.mp4", "01:00:10:00", "01:00:10:20", "00:00:00:20", "00:00:00:41"),
    # Only differs from the above by extension -- exact-string matching must
    # keep it as its own single-use clip, unlike dedup_key().
    ("Evan Fairbanks WTC.mov", "01:00:20:00", "01:00:22:00", "00:00:02:00", "00:00:02:00"),
    ("Firehouse raw.mp4", "01:00:30:00", "01:00:31:00", "00:00:01:00", "00:00:01:00"),
    ("Evan Fairbanks WTC.mp4", "01:00:05:00", "01:00:05:10", "00:00:00:10", "00:00:00:21"),
    ("Evan Fairbanks WTC.mp4", "01:00:40:00", "01:00:40:15", "00:00:00:15", "00:00:00:30"),
]


@pytest.fixture
def sample_third_parties(tmp_path):
    """A sorted workbook whose "3rd parties" sheet is the input of the FU
    grid command: real-world column layout, one clip used several times
    with its uses out of order, and a second sheet that must be ignored."""
    headers = [
        "Track", "Clip Name", "Enabled", "Sequence In", "Sequence Out",
        "Clip Duration", "Source Reel Name", "Source In", "Source Out", "Source Duration",
    ]

    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(title="Getty Videos")
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)
    ws.cell(row=2, column=2, value="GettyImages-123.mov")

    ws = wb.create_sheet(title="3rd parties")
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)
    for r, (name, seq_in, seq_out, duration, source_duration) in enumerate(THIRD_PARTIES_ROWS, start=2):
        ws.cell(row=r, column=1, value="V3")
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value="x")
        ws.cell(row=r, column=4, value=seq_in)
        ws.cell(row=r, column=5, value=seq_out)
        ws.cell(row=r, column=6, value=duration)
        ws.cell(row=r, column=7, value="3rd parties")
        ws.cell(row=r, column=10, value=source_duration)
    ws.column_dimensions["B"].width = 45

    path = tmp_path / "sorted.xlsx"
    wb.save(path)
    return path
