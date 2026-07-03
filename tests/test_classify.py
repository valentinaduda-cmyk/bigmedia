"""
Every real filename we've confirmed a category for lives here. This is the
regression net: when adding a new pattern, add its case here FIRST, then
make classify.py match it, then run pytest and make sure nothing else
flipped categories.

Do not delete a case without checking with whoever owns the sorting rules
(BigMedia project) first — each one encodes a real delivery that was
manually confirmed at some point.
"""
import pytest
from bigmedia.classify import classify

CASES = [
    # AP
    ("BM45214-3_Athens_Celebrates_Liberation_NO_SOUND_HD_50_XDCAM_50.mxf", "AP"),
    ("BMUP2804_596619_1951cafe383342bbb7631e2a71241d87.mxf", "AP"),
    ("BM3971_Olympic_Begins_Last_Trip_HD_50_XDCAM_50.mxf", "AP"),
    ("BM1215A_1677624_cb1d03443f104428b38cfb5d204b1a47.mxf", "AP"),
    ("w077854_France_Titanic_Press_Conference_HD_50_XDCAM_50.mxf", "AP"),
    ("apus1234_something.mxf", "AP"),
    ("z123456_something.mxf", "AP"),
    ("1234567_something.mxf", "AP"),
    ("28712_orig.jpg", "3rd parties"),  # looked like AP (digits_) but isn't .mxf

    # Getty
    ("GettyImages-51240904", "Getty Unknown"),
    ("GettyImages-51240904.mov", "Getty Videos"),
    ("GettyImages-51240904.jpg", "Getty Stills"),
    ("GettyImages-51240904.xyz", "Getty Unknown"),

    # Artlist
    ("33337_Heavy machinery working_By_The_Stock_Studio_Artlist_4K.mov", "Artlist"),

    # Graphics
    ("7045_old_style_graphic.mov", "Graphics"),
    ("7046_SoB_EP19_ForGrade_Full_01.mxf", "Graphics"),
    ("045_SO_EP18_01_3DExplainer_TectonicPlates_TXLS.mov", "Graphics"),

    # Camera Footage
    ("A001C003_20260101.mov", "Camera Footage"),
    ("DJI_0001.mov", "Camera Footage"),
    ("511_0498.MXF", "Camera Footage"),

    # Other agencies
    ("shutterstock_12345.mp4", "Shutterstock"),
    ("RTRWNEC_something.mxf", "Reuters"),
    ("RA-12345_bbc_clip.mov", "BBC"),

    # Fallback
    ("random_unclassified_file.mov", "3rd parties"),
]


@pytest.mark.parametrize("filename,expected", CASES, ids=[c[0] for c in CASES])
def test_classify(filename, expected):
    assert classify(filename) == expected
