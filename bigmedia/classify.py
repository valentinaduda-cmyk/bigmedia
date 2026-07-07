"""
Clip-name classification rules.

This is the ONLY file you should need to touch when a new source/pattern
shows up in a delivery. Keep it small and readable — it's a plain list of
(category, compiled regex) pairs checked in order, first match wins.

Workflow for adding a new pattern:
  1. Add the real example filename + expected category to
     tests/test_classify.py FIRST.
  2. Adjust or add a rule below.
  3. Run `pytest` and make sure every existing case still passes, not just
     the new one — categories overlap in subtle ways (see notes below).
  4. Only then consider the change done.

Category notes (why each rule looks the way it does):
  - Getty Videos / Stills / Unknown: GettyImages-<id> files are split by
    extension. "Getty Unknown" is a deliberate catch for cases where the
    extension is missing or not one we recognize as video/photo — it must
    stay listed AFTER Getty Videos/Stills so those take priority.
  - Artlist: matched by "_Artlist_" appearing anywhere in the name. Must be
    checked before AP, because Artlist filenames often start with a bare
    5-digit id that would otherwise look like an AP wire code.
  - Veritone / FOX: simple literal-prefix agencies ("VeritoneMaster_",
    "FOX_"). Checked early since the prefixes are unique enough that order
    relative to everything except Getty doesn't matter.
  - GFX: broad in-house bucket, several distinct conventions —
      * "704x_..." — legacy graphics job-code block (e.g. 7045_, 7046_)
      * "<3-4 digits>_..._TXLS.mov/mp4" — newer episode/explainer numbering,
        always ends in "_TXLS" (textless graphic element)
      * "Vi-..." — Premiere/AE placeholder-style asset names
      * "Adjustment Clip" / "Solid Color" / "Compound Clip <n>" — Premiere's
        own auto-generated track item names, matched verbatim
      * "<digits>[_<digits>]-<word>.mov", optionally with a single leading
        "<letter>_" — job-code + freeform suffix (-def, -v02, -heure,
        -sansHeure-v2, ...), e.g. "A_460-def.mov", "500-v02.mov",
        "153_160-def.mov" (episode-range pair joined by "_")
      * "<letter><digits>[<letter>]-<word>.mov" — same idea but the leading
        letter is glued directly to the digits, e.g. "P200-EN-v02.mov",
        "p300a-ENG-v03.mov"
      * "[<3-4 digits> - ]<h>h<mm> - <h>h<mm> vEN.mov" — timecode-range
        named exports, episode number prefix optional, e.g.
        "540 - 9h03 - 9h34 vEN.mov", "8h46 - 9h03 vEN.mov"
      * "ARCHIVE" anywhere in the name — archive-sourced footage re-used in
        a GFX edit, e.g. "MODIF_ARCHIVE_PLAN01.mov", "ARCHIVE-p450 v10.mov"
    This bucket is deliberately loose since it's a catch-all for in-house
    edit/graphics naming, which has no single convention.
  - Camera Footage: raw camera formats (letter+3digit+C+3digit clip names,
    DJI drone footage, P#######.MOV) plus numbered sequence formats
    "<3-4 digits>_<digits>.MXF", "CAMA_FX6_<digits>[_S###].MXF",
    "C<digits>[_S###].MP4", plus two plain markers matched anywhere in the
    name — "CANON" (camera-rig raw files, whose surrounding id format keeps
    changing) and "RMC" (production camera reference footage, e.g.
    "20260302_RMC_ITW_ENSOR_CAM1_0001.MXF") — deliberately loose for the
    same reason "BM" is loose under AP.
  - Reuters: "RTRWNEC" anywhere in the name, or an "_m<5-7 digits>_" clip id
    (e.g. "..._m909021_...") — Reuters' own internal tape/clip numbering.
  - AP: four distinct wire-code shapes —
      * "apus..." literal prefix
      * a single optional leading letter + 5-7 digits + underscore, but ONLY
        when the file is a .mxf — this is what stops generic derivative
        files like "28712_orig.jpg" from being misread as AP wire footage
      * a single optional leading letter + 5-7 digits + a DOUBLE underscore
        (e.g. "745993__HZ_Italy_..._Rewrap.mov") — same wire-code id, but
        delivered as a non-mxf rewrap. The double underscore (vs. the single
        underscore in generic derivatives like "28712_orig.jpg") is the
        signal that keeps this from swallowing unrelated files.
      * "BM..." — British Movietone / AP archive footage. Deliberately loose
        (BM + any letters/digits/hyphens + underscore) because this source's
        internal ID format has changed shape multiple times
        (BM45214-3_, BMUP2804_, BM3971_, BM1215A_1677624_, ...).
  - 3rd parties: fallback for anything that matches none of the above.
"""
import re

RULES = [
    ("Getty Videos",   re.compile(r'^GettyImages-.*\.(mov|mp4)$', re.I)),
    ("Getty Stills",   re.compile(r'^GettyImages-.*\.(jpe?g|png|tif+)$', re.I)),
    ("Getty Unknown",  re.compile(r'^GettyImages-\d+', re.I)),
    ("Veritone",       re.compile(r'^VeritoneMaster_', re.I)),
    ("FOX",            re.compile(r'^FOX_', re.I)),
    ("Artlist",        re.compile(r'_Artlist_', re.I)),
    ("Shutterstock",   re.compile(r'^shutterstock_', re.I)),
    ("Reuters",        re.compile(r'RTRWNEC|_m\d{5,7}_', re.I)),
    ("BBC",            re.compile(r'^RA-', re.I)),
    ("GFX",            re.compile(
        r'^704\d_.*'
        r'|^\d{3,4}_.*_TXLS\.(?:mov|mp4)$'
        r'|^Vi-.*'
        r'|^Adjustment Clip$'
        r'|^Solid Color$'
        r'|^Compound Clip(?:\s\d+)?$'
        r'|^(?:[A-Z]_)?\d{3,4}(?:_\d{3,4})?-[\w-]+\.mov$'
        r'|^[A-Za-z]\d{2,4}[a-z]?-[\w-]+\.mov$'
        r'|^(?:\d{3,4} - )?\d{1,2}h\d{2} - \d{1,2}h\d{2} vEN\.mov$'
        r'|ARCHIVE'
        , re.I)),
    ("Camera Footage", re.compile(
        r'^[A-Z]\d{3}C\d{3}_'
        r'|^DJI'
        r'|^P\d{7}\.MOV'
        r'|^\d{3,4}_\d{3,6}\.MXF$'
        r'|^CAMA_FX6_\d+(?:_S\d+)?\.MXF$'
        r'|^C\d{3,5}(?:_S\d+)?\.MP4$'
        r'|CANON'
        r'|RMC'
        , re.I)),
    ("AP",             re.compile(r'^(?:apus|[a-z]?\d{5,7}_.*\.mxf$|[a-z]?\d{5,7}__|BM[A-Za-z0-9-]*_)', re.I)),
]

CATEGORY_ORDER = [
    "AP", "Getty Videos", "Getty Stills", "Getty Unknown", "Reuters",
    "Shutterstock", "Artlist", "GFX", "Camera Footage", "BBC",
    "Veritone", "FOX", "3rd parties",
]

DEFAULT = "3rd parties"


def classify(name: str) -> str:
    """Return the category for a clip filename. First matching rule wins."""
    name = (name or "").strip()
    for category, pattern in RULES:
        if pattern.search(name):
            return category
    return DEFAULT
