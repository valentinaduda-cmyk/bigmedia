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
    ("745993__HZ_Italy_Legionnaries_1080i50-xdcam-50m_Rewrap.mov", "AP"),  # non-mxf rewrap, double underscore
    ("28712_orig.jpg", "3rd parties"),  # looked like AP (digits_) but isn't .mxf, and single underscore

    # Getty
    ("GettyImages-51240904", "Getty Unknown"),
    ("GettyImages-51240904.mov", "Getty Videos"),
    ("GettyImages-51240904.jpg", "Getty Stills"),
    ("GettyImages-51240904.xyz", "Getty Unknown"),

    # Artlist
    ("33337_Heavy machinery working_By_The_Stock_Studio_Artlist_4K.mov", "Artlist"),

    # GFX
    ("7045_old_style_graphic.mov", "GFX"),
    ("7046_SoB_EP19_ForGrade_Full_01.mxf", "GFX"),
    ("045_SO_EP18_01_3DExplainer_TectonicPlates_TXLS.mov", "GFX"),

    # Camera Footage
    ("A001C003_20260101.mov", "Camera Footage"),
    ("DJI_0001.mov", "Camera Footage"),
    ("511_0498.MXF", "Camera Footage"),
    ("CAMA_FX6_3630.MXF", "Camera Footage"),
    ("C2000.MP4", "Camera Footage"),
    ("B_0001C001A260205_0042337W_CANON_S004.MXF", "Camera Footage"),
    ("CAMA_FX6_3687_S009.MXF", "Camera Footage"),
    ("C2012_S005.MP4", "Camera Footage"),
    ("C2015_S008.MP4", "Camera Footage"),
    ("CAMA_FX6_3671.MXF", "Camera Footage"),
    ("C2008.MP4", "Camera Footage"),
    ("B_0001C001A260205_0042337W_CANON.MXF", "Camera Footage"),
    ("20260203_RMC_ITW_ENSOR_CAM2_0001_S008.MP4", "Camera Footage"),
    ("A_0001C001A260204_113406T5_CANON_S003.MXF", "Camera Footage"),
    ("A_0001C001A260204_113406T5_CANON.MXF", "Camera Footage"),
    ("A_0001C001A260204_113406T5_CANON_S000.MXF", "Camera Footage"),
    ("A_0001C001A260204_113406T5_CANON_S002.MXF", "Camera Footage"),
    ("20260302_RMC_ITW_ENSOR_CAM1_0001.MXF", "Camera Footage"),
    ("20260302_RMC_ITW_ENSOR_CAM1_0001_S000.MXF", "Camera Footage"),
    ("20260302_RMC_ITW_ENSOR_CAM1_0001_S001.MXF", "Camera Footage"),

    # GFX (in-house edit/graphics project names)
    ("Vi-SYNTHE-MarcyMcGinnis-DEF.mov", "GFX"),
    ("Adjustment Clip", "GFX"),
    ("Solid Color", "GFX"),
    ("Vi-SYNTHE-StuartAllan-DEF.mov", "GFX"),
    ("A_460-def.mov", "GFX"),
    ("470-def.mov", "GFX"),
    ("500-v02.mov", "GFX"),
    ("Compound Clip 5", "GFX"),
    ("540 - 9h03 - 9h34 vEN.mov", "GFX"),
    ("Vi-P550_551.mov", "GFX"),
    ("552-sansHeure-v2.mov", "GFX"),
    ("Vi-SYNTHE-LindaRoth-DEF.mov", "GFX"),
    ("A_571-def.mov", "GFX"),
    ("Vi-10h28.mov", "GFX"),
    ("580 - 10h28 - 9h59 vEN.mov", "GFX"),
    ("151-Vdef_fixe.mov", "GFX"),
    ("Vi-100393.mov", "GFX"),
    ("Vi-TITRE01.mov", "GFX"),
    ("153_160-def.mov", "GFX"),
    ("153_160-heure.mov", "GFX"),
    ("MODIF_ARCHIVE_PLAN01.mov", "GFX"),
    ("MODIF_ARCHIVE_PLAN02_b.mov", "GFX"),
    ("P200-EN-v02.mov", "GFX"),
    ("p300a-ENG-v03.mov", "GFX"),
    ("ARCHIVE-p450 v10.mov", "GFX"),
    ("8h46 - 9h03 vEN.mov", "GFX"),
    ("ARCHIVE-651-v2.mov", "GFX"),
    ("ARCHIVE-p653-v2.mov", "GFX"),
    ("ARCHIVE-p654-v2.mov", "GFX"),

    # Other agencies
    ("shutterstock_12345.mp4", "Shutterstock"),
    ("RTRWNEC_something.mxf", "Reuters"),
    ("RA-12345_bbc_clip.mov", "BBC"),
    ("Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE ONE - 9_11 compilation_m909021_0-14-39-0-15-37 (1).mp4", "Reuters"),
    ("Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE TWO - 9_11 compilation_m1241025_0-35-40-0-36-00.mov", "Reuters"),
    ("VeritoneMaster_2B12018_00123868 (3).mov", "Veritone"),
    ("FOX_2001-09-11_BNYC_NY5178_004_WTCATTACKAFTERMATHANDEVACUATIONS_00152815.mov", "FOX"),
    ("FOX_2001-09-11_BNYC_NY5178_004_WTCATTACKAFTERMATHANDEVACUATIONS_00151717.mov", "FOX"),

    # Fallback
    ("random_unclassified_file.mov", "3rd parties"),
]


@pytest.mark.parametrize("filename,expected", CASES, ids=[c[0] for c in CASES])
def test_classify(filename, expected):
    assert classify(filename) == expected
