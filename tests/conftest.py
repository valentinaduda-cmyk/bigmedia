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

    headers = ["Clip Name", "Duration", "Notes"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="CCCCCC")

    rows = [
        ("BM45214-3_Athens_Celebrates_Liberation_NO_SOUND_HD_50_XDCAM_50.mxf", "00:01:20:00", "AP archive"),
        ("045_SO_EP18_01_3DExplainer_TectonicPlates_TXLS.mov", "00:00:45:00", "graphics element"),
        ("GettyImages-51240904", "", "unknown media type"),
        ("GettyImages-99999999.mov", "00:02:00:00", "getty video"),
        ("shutterstock_777.mp4", "00:00:10:00", ""),
        ("random_file_xyz.mov", "00:00:05:00", "unclassified"),
        ("BM45214-3_Athens_Celebrates_Liberation_NO_SOUND_HD_50_XDCAM_50 (1).mxf", "00:01:20:00", "duplicate of row 2"),
    ]
    for r, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=val if val != "" else None)
        ws.cell(row=r, column=2).number_format = "hh:mm:ss"

    for col_letter, width in [("A", 60), ("B", 15), ("C", 25)]:
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"

    path = tmp_path / "sample_master.xlsx"
    wb.save(path)
    return path
