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
  - Graphics: two distinct in-house numbering conventions —
      * "704x_..." — legacy graphics job-code block (e.g. 7045_, 7046_)
      * "<3-4 digits>_..._TXLS.mov/mp4" — newer episode/explainer numbering,
        always ends in "_TXLS" (textless graphic element)
  - Camera Footage: raw camera formats (letter+3digit+C+3digit clip names,
    DJI drone footage, P#######.MOV) plus a numbered "<3-4 digits>_<digits>.MXF"
    sequence format.
  - AP: three distinct wire-code shapes —
      * "apus..." literal prefix
      * a single optional leading letter + 5-7 digits + underscore, but ONLY
        when the file is a .mxf — this is what stops generic derivative
        files like "28712_orig.jpg" from being misread as AP wire footage
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
    ("Artlist",        re.compile(r'_Artlist_', re.I)),
    ("Shutterstock",   re.compile(r'^shutterstock_', re.I)),
    ("Reuters",        re.compile(r'RTRWNEC', re.I)),
    ("BBC",            re.compile(r'^RA-', re.I)),
    ("Graphics",       re.compile(r'^(?:704\d_.*|\d{3,4}_.*_TXLS\.(?:mov|mp4)$)', re.I)),
    ("Camera Footage", re.compile(r'^[A-Z]\d{3}C\d{3}_|^DJI|^P\d{7}\.MOV|^\d{3,4}_\d{3,6}\.MXF$', re.I)),
    ("AP",             re.compile(r'^(?:apus|[a-z]?\d{5,7}_.*\.mxf$|BM[A-Za-z0-9-]*_)', re.I)),
]

CATEGORY_ORDER = [
    "AP", "Getty Videos", "Getty Stills", "Getty Unknown", "Reuters",
    "Shutterstock", "Artlist", "Graphics", "Camera Footage", "BBC",
    "3rd parties",
]

DEFAULT = "3rd parties"


def classify(name: str) -> str:
    """Return the category for a clip filename. First matching rule wins."""
    name = (name or "").strip()
    for category, pattern in RULES:
        if pattern.search(name):
            return category
    return DEFAULT
