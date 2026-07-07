# bigmedia — context for Claude Code

This repo automates recurring Excel work for KM Records' BigMedia
workflow: sorting media-master clip lists by source (AP, Getty, Reuters,
Shutterstock, Artlist, in-house GFX, Camera Footage, BBC, and a
"3rd parties" fallback), deduping, and eventually timecode-based
calculations. It replaced a set of loose one-off scripts that had no
tests and no history of *why* each rule existed.

## What this is for

Valentina (or someone on her team) will hand you an .xlsx master list and
ask for an operation — most often "sort this" or "add this new pattern to
the sorter." Sometimes they'll give you a specific filename and the
category it should land in, or should NOT land in (a false positive).

## The one rule that matters most

**`bigmedia/classify.py` is the single source of truth for sorting logic.**
Never duplicate its RULES elsewhere. When asked to add or fix a pattern:

1. Add the real example filename + expected category as a new case in
   `tests/test_classify.py` first.
2. Adjust `RULES` in `classify.py` — read the docstring at the top of that
   file first, it explains why each rule is shaped the way it is (order
   matters: some categories must be checked before others to avoid false
   positives, e.g. Artlist before AP, Getty Videos/Stills before Getty
   Unknown).
3. Run `pytest` and confirm every case passes, not just the new one.
   Categories have overlapping-looking patterns (digit prefixes especially)
   — a fix for one filename has broken previously-confirmed ones before.
4. Only then report the change as done. Summarize which categories'
   row counts changed and by how much when re-run against a real file, so
   it's easy to sanity-check nothing moved unexpectedly.

Do not delete or "clean up" existing test cases without checking first —
each one encodes a real delivery someone manually confirmed.

## Conventions

- Business logic (classify, dedupe key, timecode math) lives in its own
  module with no I/O in it. openpyxl plumbing lives in `xlsx_utils.py` and
  is shared across commands — don't reimplement "copy this sheet's
  styling" inline again.
- Every workbook-producing command writes a `"Worksheet"` sheet first: a
  copy of the input, as a backup. Preserve this pattern in new commands
  unless explicitly told otherwise.
- The input is expected to already have a "Source Reel Name" column. Every
  output sheet (including the "Worksheet" backup) overwrites that column
  per row with the name of whichever sheet the row ended up on — e.g. rows
  sorted into "AP" get "AP", rows kept by dedupe get "Deduped". This means
  "Worksheet" is no longer byte-identical to the input; only that one
  column differs.
- CLI commands are thin wrappers in `cli.py` around a plain function
  (`sort_workbook(src, out, ...)`, `dedupe_workbook(src, out, ...)`) that
  returns a small dict of counts. Follow this shape for new operations so
  they're testable without going through argparse.
- Real client files never get committed. `data/in/` and `data/out/` are
  gitignored; tests use the synthetic fixture in `tests/conftest.py`.

## Known category shapes (as of the last update)

- **AP**: `apus...` literal prefix; a single optional leading letter +
  5-7 digits + underscore but only for `.mxf` files (bare digit-prefixed
  non-mxf files, e.g. `28712_orig.jpg`, are NOT AP); `BM...` — British
  Movietone/AP archive, format has changed shape several times
  (`BM45214-3_`, `BMUP2804_`, `BM3971_`, `BM1215A_1677624_`), so the rule
  is deliberately loose (`BM` + letters/digits/hyphens + `_`).
- **Getty Videos / Stills / Unknown**: split by extension;
  "Getty Unknown" catches missing/unrecognized extensions and must be
  checked after the other two.
- **Artlist**: `_Artlist_` anywhere in the name; checked before AP since
  Artlist names often start with a bare digit id.
- **GFX**: two conventions — legacy `704x_...` job-code block, and
  newer `<3-4 digits>_..._TXLS.mov/mp4` episode numbering.
- **Camera Footage**: raw camera formats (`[A-Z]###C###_`, `DJI`,
  `P#######.MOV`) plus `<3-4 digits>_<digits>.MXF`.
- **3rd parties**: fallback, matches nothing else.

When in doubt about a new pattern's edge cases, ask rather than guessing —
false positives here (e.g. treating a generic file as AP) are worse than
leaving something in "3rd parties" for manual review.

## Web UI (planned)

CLI-only today; a Web UI for the internal team (shared server) is planned
next. When wireframing/quicksketching that UI, use Excalidraw.
