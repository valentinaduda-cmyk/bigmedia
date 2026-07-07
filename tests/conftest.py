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
