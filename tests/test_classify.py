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
    ("4079167_PERSIAN_GULF_EXERCISES_1080I50-XDCAM-50M", "AP"),
    ("209317_ITALY_NATO_LAUNCHES_DOGFISH_2001_EXERCISE_HD_50_XDCAM_50", "AP"),
    ("865672_CHINA_ISLAND_LANDING_DRILL_1080I50-XDCAM-50M", "AP"),
    ("CCTV016909_CHINA_MILITARY_ASSESSMENT", "AP"),
    ("CCTV024952_ARGENTINA_MISSING_SUBMARINESEARCHUPDATE_1080I50-XDCAM-50M", "AP"),
    ("4030219_4030219_1080I50-XDCAM-50M", "AP"),
    ("CCTV040914_CHINA_PLAANNIVERSARY", "AP"),
    ("901631_INDIA_US_1080I50-XDCAM-50M", "AP"),
    ("2026675_SERBIA_PIPELINE_1080I50-XDCAM-50M", "AP"),
    ("2051613_AT_SEA_CARRIERS_1080I50-XDCAM-50M-CORRUPT", "AP"),
    ("4501071_Cuba_Russian_Submarine_1080i50-xdcam-50m", "AP"),
    ("G03811_TOMAHAWK_HD_50_XDCAM_50", "AP"),
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
    ("7038_AD_DINO_3DANIM_07_2_HEAD_V1.MP4", "GFX"),
    ("7038_AD_DINO_3DANIM_07_3_SKULL_V1.MP4", "GFX"),
    ("7038_AD_DINO_MAP_10_3_TUGRUGIINSHIREE_V1.MP4", "GFX"),

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
    # Sony-style clip id behind a shoot-date prefix -- same "C<digits>.MP4"
    # camera convention as the bare "C2000.MP4" cases above.
    ("20230515_C0578.MP4", "Camera Footage"),
    ("20230515_C0578_S002.MP4", "Camera Footage"),
    # "RMC" is a delimited production marker, not a substring: it must not
    # fire inside an ordinary word ("WinsoRMCay").
    ("LibertyBond-WinsorMcCay.jpg", "3rd parties"),

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
    ("Rich", "GFX"),
    ("Fusion Title", "GFX"),
    ("Richard", "3rd parties"),  # "Rich" must match as a whole word only, not a substring
    ("MODIF_ARCHIVE_PLAN01.mov", "GFX"),
    ("MODIF_ARCHIVE_PLAN02_b.mov", "GFX"),
    ("P200-EN-v02.mov", "GFX"),
    ("p300a-ENG-v03.mov", "GFX"),
    ("ARCHIVE-p450 v10.mov", "GFX"),
    ("8h46 - 9h03 vEN.mov", "GFX"),
    ("ARCHIVE-651-v2.mov", "GFX"),
    ("ARCHIVE-p653-v2.mov", "GFX"),
    ("ARCHIVE-p654-v2.mov", "GFX"),

    # Content Mint
    ("CM0101-OCST-0102610.mov", "Content Mint"),
    ("CM0103-PDS-GT-Hippo-Group-Juvenile-Baby-Resting-Pilanesberg-SA-90s-UHD-ProResHQ.mov", "Content Mint"),
    ("CM0001-AW-0004294.mov", "Content Mint"),
    ("CM0100-CMP-180813-DALE0034.MXF", "Content Mint"),
    ("CM0001-AW-0004287.mov", "Content Mint"),
    ("CM0101-OCST-0100035", "Content Mint"),

    # Other agencies
    ("AdobeStock_517623332_Video_HD_Preview.mov", "AdobeStock"),  # has "Preview" -- must not fall into Shutterstock
    ("shutterstock_12345.mp4", "Shutterstock"),
    ("1106944703-PREVIEW.MP4", "Shutterstock"),
    ("3509859073-", "Shutterstock"),
    ("3728349849-", "Shutterstock"),
    ("3759691947-", "Shutterstock"),
    ("3604015041-", "Shutterstock"),
    ("HDS015570_PROXY.MP4", "BBC"),
    ("BRD071397_PROXY.MP4", "BBC"),
    ("RTRWNEC_something.mxf", "Reuters"),
    ("RA-12345_bbc_clip.mov", "BBC"),
    ("Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE ONE - 9_11 compilation_m909021_0-14-39-0-15-37 (1).mp4", "Reuters"),
    ("Clip from USA_ NEW YORK_ WORLD TRADE CENTER RUSHES COMPILATION TAPE TWO - 9_11 compilation_m1241025_0-35-40-0-36-00.mov", "Reuters"),
    ("VeritoneMaster_2B12018_00123868 (3).mov", "Veritone"),
    ("FOX_2001-09-11_BNYC_NY5178_004_WTCATTACKAFTERMATHANDEVACUATIONS_00152815.mov", "FOX"),
    ("FOX_2001-09-11_BNYC_NY5178_004_WTCATTACKAFTERMATHANDEVACUATIONS_00151717.mov", "FOX"),

    # Reuters -- "_m<5-7 digits>" no longer needs a trailing underscore
    ("Clip from SYRIA compilation_m1241025.mov", "Reuters"),
    ("tape_M909021-0-14-39.mp4", "Reuters"),

    # GFX -- 3D template / textless markers anywhere in the name
    ("SO_EP21_3DTransition_Volcano_v03.mov", "GFX"),
    ("SO_EP21_3dexplainer_Tectonics.mp4", "GFX"),
    ("EP12_intro_txls.mp4", "GFX"),

    # Getty -- "gettyimages" anywhere, not only as a prefix
    ("KM_gettyimages-51240904_master.mov", "Getty Videos"),
    ("KM_GettyImages-51240904_master.jpg", "Getty Stills"),
    ("KM_gettyimages-51240904_master.xyz", "Getty Unknown"),

    # AP -- wire format markers anywhere in the name
    ("ITALY_NATO_DRILL_HD_50_XDCAM_50.mov", "AP"),
    ("SPAIN_FLOODS_1080i50-50m.mov", "AP"),

    # Shutterstock "Preview" is lowest priority -- other categories win first
    ("some_stock_clip_Preview.mov", "Shutterstock"),
    ("SO_EP21_3DTransition_Volcano_Preview.mov", "GFX"),
    ("KM_gettyimages-51240904_Preview.mov", "Getty Videos"),

    # "ARCHIVE" is a GFX marker only when it's the in-house all-caps token on
    # a name that doesn't start with a wire id. AP slugs carry "Archive" as an
    # ordinary word inside the story description and must stay AP -- GFX is
    # checked before AP, so a loose match here silently steals wire footage.
    ("MODIF_ARCHIVE_PLAN01.mov", "GFX"),
    ("ARCHIVE-p450 v10.mov", "GFX"),
    ("4465876_Archive_UK_Greece_Parthenon_Marbles_1080i50-xdcam-50m.mxf", "AP"),
    ("4353674_Archive_F35_1080i50-xdcam-50m", "AP"),
    ("4435551_Archive_F_16s", "AP"),
    ("4211299_ARCHIVE_Pei_2_1080i50-xdcam-50m.mxf", "AP"),
    ("4611760_Archive_France_Louvre_Gallery_HD_50_XDCAM_50.mxf", "AP"),
    ("4257234_Vatican_Pius_XII_Archive_1080i50-xdcam-50m.mxf", "AP"),
    ("730277_Italy_Vatican_Archives_HD_50_XDCAM_50.mxf", "AP"),
    ("4359193_ARCHIVE_Dan_Brown_1080i50-xdcam-50m.mxf", "AP"),
    # A YouTube rip whose title happens to end in the English word "Archive".
    ("Donald Menzel 1973 Eclipse - Red Panda Koala UFO History Archive (360p, h264).mp4", "3rd parties"),

    # Fallback
    ("random_unclassified_file.mov", "3rd parties"),
    (12345.0, "3rd parties"),
]


@pytest.mark.parametrize("filename,expected", CASES, ids=[c[0] for c in CASES])
def test_classify(filename, expected):
    assert classify(filename) == expected
