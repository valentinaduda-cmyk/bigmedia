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
  - Getty Videos / Stills / Unknown: "gettyimages" anywhere in the name
    (case-insensitive), split by extension. "Getty Unknown" is a deliberate
    catch for cases where the extension is missing or not one we recognize
    as video/photo — it must stay listed AFTER Getty Videos/Stills so those
    take priority.
  - AdobeStock: "AdobeStock_" literal prefix.
  - Shutterstock: "shutterstock_" literal prefix or a 10+ digit id followed
    by "-" (e.g. "1106944703-PREVIEW.MP4"). Plus a LATE rule: "PREVIEW"
    anywhere in the name, checked only after every other category has
    failed — see LATE_RULES below. Many sources ship "_Preview" exports
    (AdobeStock, GFX 3D templates, Getty masters), so it is the weakest
    signal in the file and must never outrank a real source match.
  - BBC: "RA-" literal prefix, or "PROXY" anywhere in the name
    (case-insensitive) -- BBC low-res proxy exports, e.g.
    "HDS015570_PROXY.MP4", "BRD071397_PROXY.MP4".
  - Artlist: matched by "_Artlist_" appearing anywhere in the name. Must be
    checked before AP, because Artlist filenames often start with a bare
    5-digit id that would otherwise look like an AP wire code.
  - Veritone / FOX / Content Mint: simple literal-prefix agencies
    ("VeritoneMaster_", "FOX_", "CM<4 digits>-"). Checked early since the
    prefixes are unique enough that order relative to everything except
    Getty doesn't matter.
  - GFX: broad in-house bucket, several distinct conventions —
      * "70xx_..." — legacy graphics job-code block, any 4-digit code in
        the 7000s (e.g. 7045_, 7046_, 7038_)
      * "TXLS" anywhere — textless graphic element, newer episode/explainer
        numbering (e.g. "045_SO_EP18_01_..._TXLS.mov")
      * "3DTransition" / "3DExplainer" anywhere — 3D template elements
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
      * all-caps "ARCHIVE" on a name that doesn't start with a digit —
        archive-sourced footage re-used in a GFX edit, e.g.
        "MODIF_ARCHIVE_PLAN01.mov", "ARCHIVE-p450 v10.mov". This one is
        matched case-SENSITIVELY, by its own rule entry below, and both
        halves of that restriction are load-bearing: GFX is checked before
        AP, so a loose "archive" match steals AP wire footage whose story
        slug happens to contain the word ("4465876_Archive_UK_Greece_
        Parthenon_Marbles_1080i50-xdcam-50m.mxf" is AP, not GFX). Requiring
        a non-digit first character rejects the wire-id prefix; requiring
        all caps rejects the English word inside a prose filename (a
        YouTube rip titled "...UFO History Archive (360p).mp4").
      * "Rich" / "Fusion Title" — Fusion/Resolve auto-generated track item
        names, same idea as "Adjustment Clip"/"Solid Color" above but
        matched as a whole word anywhere in the name rather than only
        verbatim, so a numbered/suffixed variant (e.g. "Rich_02.mov")
        still lands here. Word-boundary matched specifically because
        "Rich" is short enough to otherwise false-positive inside an
        unrelated word/name (e.g. "Richard").
    This bucket is deliberately loose since it's a catch-all for in-house
    edit/graphics naming, which has no single convention.
  - Camera Footage: raw camera formats (letter+3digit+C+3digit clip names,
    DJI drone footage, P#######.MOV) plus numbered sequence formats
    "<3-4 digits>_<digits>.MXF", "CAMA_FX6_<digits>[_S###].MXF",
    "[<prefix>_]C<digits>[_S###].MP4" — the Sony-style clip id, either
    standing alone ("C2000.MP4") or behind a shoot-date//reel prefix
    ("20230515_C0578.MP4"), which is why the "C" is anchored to the start of
    the name OR to an underscore rather than allowed anywhere (a free-
    floating "C<digits>" would swallow unrelated names) — plus two plain
    markers matched anywhere in the name: "CANON" (camera-rig raw files,
    whose surrounding id format keeps changing) and "RMC" (production camera
    reference footage, e.g. "20260302_RMC_ITW_ENSOR_CAM1_0001.MXF").
    "RMC" is short enough to appear inside an ordinary word, so it is
    matched only when not flanked by letters — without that, a stills file
    like "LibertyBond-WinsorMcCay.jpg" matches on the "rMc" of "WinsorMcCay"
    and gets filed as camera rushes.
  - Reuters: "RTRWNEC" anywhere in the name, or an "_m<5-7 digits>_" clip id
    (e.g. "..._m909021_...", "..._m1241025.mov" — no trailing underscore
    required) — Reuters' own internal tape/clip numbering.
  - AP: "XDCAM" or "1080I50" anywhere in the name — AP wire delivery format
    markers — plus four distinct wire-code shapes —
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
    ("Getty Videos",   re.compile(r'gettyimages.*\.(mov|mp4)$', re.I)),
    ("Getty Stills",   re.compile(r'gettyimages.*\.(jpe?g|png|tif+)$', re.I)),
    ("Getty Unknown",  re.compile(r'gettyimages', re.I)),
    ("Veritone",       re.compile(r'^VeritoneMaster_', re.I)),
    ("FOX",            re.compile(r'^FOX_', re.I)),
    ("Content Mint",   re.compile(r'^CM\d{4}-', re.I)),
    ("Artlist",        re.compile(r'_Artlist_', re.I)),
    ("AdobeStock",     re.compile(r'^AdobeStock_', re.I)),
    ("Shutterstock",   re.compile(r'^(?:shutterstock_|\d{10,}-)', re.I)),
    ("Reuters",        re.compile(r'RTRWNEC|_m\d{5,7}', re.I)),
    ("BBC",            re.compile(r'^RA-|PROXY', re.I)),
    ("GFX",            re.compile(
        r'^70\d{2}_.*'
        r'|TXLS'
        r'|3DTransition'
        r'|3DExplainer'
        r'|^Vi-.*'
        r'|^Adjustment Clip$'
        r'|^Solid Color$'
        r'|^Compound Clip(?:\s\d+)?$'
        r'|^(?:[A-Z]_)?\d{3,4}(?:_\d{3,4})?-[\w-]+\.mov$'
        r'|^[A-Za-z]\d{2,4}[a-z]?-[\w-]+\.mov$'
        r'|^(?:\d{3,4} - )?\d{1,2}h\d{2} - \d{1,2}h\d{2} vEN\.mov$'
        r'|\bRich\b'
        r'|\bFusion Title\b'
        , re.I)),
    # Deliberately NOT re.I, and deliberately its own entry so the rest of
    # GFX keeps matching case-insensitively. See the ARCHIVE note above.
    ("GFX",            re.compile(r'^(?!\d).*ARCHIVE')),
    ("Camera Footage", re.compile(
        r'^[A-Z]\d{3}C\d{3}_'
        r'|^DJI'
        r'|^P\d{7}\.MOV'
        r'|^\d{3,4}_\d{3,6}\.MXF$'
        r'|^CAMA_FX6_\d+(?:_S\d+)?\.MXF$'
        r'|(?:^|_)C\d{3,5}(?:_S\d+)?\.MP4$'
        r'|CANON'
        r'|(?<![A-Za-z])RMC(?![A-Za-z])'
        , re.I)),
    ("AP",             re.compile(
        r'XDCAM'
        r'|1080I50'
        r'|^(?:'
        r'apus|'
        r'[A-Za-z]?\d{5,7}\.mxf$|'
        r'[A-Za-z]?\d{5,7}__|'
        r'[A-Za-z]?\d{5,7}_(?:[A-Z][A-Za-z0-9_\-]*|\d{5,10})|'
        r'cctv\d+|'
        r'G\d{5,}_|'
        r'BM[A-Za-z0-9-]*_' 
        r')', re.I
    )),
]

# Checked only after every rule in RULES has failed. "PREVIEW" is a weak
# signal -- lots of sources (AdobeStock, GFX 3D templates, Getty masters)
# ship "..._Preview.mov" exports -- so it may only claim a file that no
# stronger rule wanted.
LATE_RULES = [
    ("Shutterstock",   re.compile(r'PREVIEW', re.I)),
]

CATEGORY_ORDER = [
    "AP", "Getty Videos", "Getty Stills", "Getty Unknown", "Reuters",
    "Shutterstock", "AdobeStock", "Artlist", "GFX", "Camera Footage", "BBC",
    "Veritone", "FOX", "Content Mint", "3rd parties",
]

DEFAULT = "3rd parties"


def classify(name: str) -> str:
    """Return the category for a clip filename. First matching rule wins."""
    if name is None:
        text = ""
    elif isinstance(name, str):
        text = name.strip()
    else:
        text = str(name).strip()

    # Generic numeric derivative files like "28712_orig.jpg" are *not* AP
    # wire footage even though they begin with a 5-7 digit prefix.
    if re.fullmatch(r"\d{5,7}_[A-Za-z][A-Za-z0-9_-]*\.(?:jpe?g|png|tif+)", text, re.I):
        return DEFAULT

    for category, pattern in RULES + LATE_RULES:
        if pattern.search(text):
            return category
    return DEFAULT
