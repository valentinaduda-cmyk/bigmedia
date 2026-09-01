# bigmedia

Excel sheet data processing tool for KM Records' BigMedia workflow. Takes a
media-master clip list (`.xlsx`) and automates the recurring spreadsheet
work around it:

- **sort** — classify every clip by source (AP, Getty, Reuters,
  Shutterstock, Artlist, in-house GFX, Camera Footage, BBC, or a "3rd
  parties" fallback) into separate sheets
- **dedupe** — split a list into deduped rows + a "Duplicates" sheet, so
  nothing disappears silently
- **group** — second pass over a sorted workbook: groups repeat
  occurrences of the same clip together and totals their duration
  (frame/timecode-accurate)

Currently a local CLI, run by hand or via Claude Code (see `CLAUDE.md` for
the conventions this repo follows).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
```

## Usage

Drop input files in `data/in/` (gitignored — never commit real client
data), then:

```bash
bigmedia sort data/in/EP21_master.xlsx -o data/out/EP21_sorted.xlsx
bigmedia dedupe data/in/EP21_master.xlsx -o data/out/EP21_deduped.xlsx
bigmedia group data/out/EP21_sorted.xlsx -o data/out/EP21_grouped.xlsx
bigmedia fu-grid data/out/EP21_sorted.xlsx -o data/out/EP21_fu_grid.xlsx
bigmedia compare data/in/EP21_july.xlsx data/out/EP21_sorted.xlsx
bigmedia fix-getty data/in/OLD_VERSIONS data/out/sorted -o data/out/getty_fixed
```

`fix-getty` repairs the Getty Videos/Getty Stills split of freshly sorted
workbooks against older, hand-checked versions of the same episodes. Some
EDL exports write Getty stills as bare ids with no `.jpg` and with
`Source` set to "Getty Images - Footage", so `sort` has nothing to go on
and files them as videos; the old hand-corrected delivery is the only
record of which ids are stills. Old and new files are paired by EPISODE
NUMBER (the "Episode 20" value in the sorted file's first column, or an
"EP20" in the old file's name) — filenames don't match across versions.

Only the two Getty sheets are touched, only the sheet a row sits on
changes (no cell value is rewritten, including `Source`), and clips the
old version doesn't have at all are left where `sort` put them and listed
as unmatched. A `Getty Stills` sheet is created if the new file has none.
The output is re-grouped automatically (`--no-group` to skip), so it is a
drop-in replacement for a `group` output.

`compare` diffs two versions of the same episode category by category:
how many clips each category holds on either side, which clips the new
version gained, which it lost, and which merely changed category. Those
last two are kept apart deliberately — a clip that moved from
"3rd parties" to "Getty Videos" is not new material and must not be
cleared twice. Output is a `Summary` sheet (`Clips in old`, `Clips in
new`, `Added`, `Removed`, `Moved in`, `Moved out`, plus a TOTAL row) then
`<category> added` / `<category> removed` sheets carrying the real rows
behind a `Status` column ("added", "moved from Reuters", "removed",
"moved to AP").

Point it at two folders and every episode present in both is compared,
paired by episode number; one report per episode lands in
`<new folder>/compared` (or `-o`). Categories pair by sheet name
case-insensitively, with `Getty pics`/`Getty Stills` and `graphics`/`GFX`
aliased; backup tabs ("Worksheet", "... Master XML", "Kopie listu ...")
are not categories but still count as evidence a clip exists in that file.
Filenames match via the same `dedup_key` as `dedupe`, case-insensitively
(`--case-sensitive` to turn that off); `--unique` lists each clip once
instead of one row per use.

`fix-getty` repairs the Getty Videos/Getty Stills split of freshly sorted
workbooks against older, hand-checked versions of the same episodes. Some
EDL exports write Getty stills as bare ids with no `.jpg` and with
`Source` set to "Getty Images - Footage", so `sort` has nothing to go on
and files them as videos; the old hand-corrected delivery is the only
record of which ids are stills. Old and new files are paired by EPISODE
NUMBER (the "Episode 20" value in the sorted file's first column, or an
"EP20" in the old file's name) — filenames don't match across versions.

Only the two Getty sheets are touched, only the sheet a row sits on
changes (no cell value is rewritten, including `Source`), and clips the
old version doesn't have at all are left where `sort` put them and listed
as unmatched. A `Getty Stills` sheet is created if the new file has none.
The output is re-grouped automatically (`--no-group` to skip), so it is a
drop-in replacement for a `group` output.

`fu-grid` builds the legal follow-up grid from a sorted workbook's "3rd
parties" sheet: one row per clip (matched by exact clip name), its
`TOTAL USES` and the summed duration of those uses, plus empty columns
for the legal team. The source sheet is copied through untouched as the
output's first sheet. Takes `--sheet` (default "3rd parties"),
`--duration-column` and `--fps` as well.

All three default to `<input>_sorted.xlsx` / `<input>_deduped.xlsx` /
`<input>_grouped.xlsx` next to the input file if `-o` is omitted. Use
`--name-column` if a workbook's filename column isn't called "Clip Name";
`group` also takes `--duration-column` (default "Clip Duration") and
`--fps` (default 25).

`<input>` can also be a folder — every `.xlsx` file in it gets processed,
each writing its own `<name>_sorted.xlsx` / `<name>_deduped.xlsx` /
`<name>_grouped.xlsx`:

```bash
bigmedia sort data/in -o data/out
```

If `-o` is omitted for a folder input, outputs are written alongside each
input file instead.

## Running tests

```bash
pytest
```

Do this before and after any change to `bigmedia/classify.py` — the test
suite in `tests/test_classify.py` is the record of every filename pattern
we've confirmed a category for. See the module docstring in `classify.py`
for the reasoning behind each rule.

## Layout

```
bigmedia/
  cli.py               single entrypoint, one subcommand per operation
  classify.py          sorting rules — the file you'll touch most often
  sort_workbook.py      "sort" command logic
  dedupe.py               "dedupe" command logic
  group_duplicates.py       "group" command logic
  fu_grid.py                  "fu-grid" command logic
  compare_versions.py           "compare" command logic (old vs new version)
  getty_split.py                  "fix-getty" command logic (Getty stills
                                    misfiled as videos)
  timecode.py                   frame/timecode conversions (shared helper,
                                  used by group_duplicates and fu_grid)
  xlsx_utils.py                  shared openpyxl helpers (styling, sheet
                                   copying)
tests/                      pytest suite + a synthetic fixture generator
data/in/, data/out/          working directories, gitignored
```

## Adding a new operation

Add a function to a new or existing module under `bigmedia/`, wire a
subcommand for it in `cli.py`, and add tests. Keep the same shape as
`sort_workbook()` / `dedupe_workbook()`: take an input path and an output
path, return a small dict of counts/results the CLI can print.
