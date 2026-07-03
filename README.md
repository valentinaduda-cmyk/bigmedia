# bigmedia

Excel automation for media-master workflows: sorting clip lists by source,
deduping, and (soon) timecode-based calculations. Built to be driven by
Claude Code — see `CLAUDE.md` for the conventions it follows in this repo.

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
```

Both default to `<input>_sorted.xlsx` / `<input>_deduped.xlsx` next to the
input file if `-o` is omitted. Use `--name-column` if a workbook's filename
column isn't called "Clip Name".

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
  cli.py             single entrypoint, one subcommand per operation
  classify.py         sorting rules — the file you'll touch most often
  sort_workbook.py     "sort" command logic
  dedupe.py             "dedupe" command logic
  timecode.py            frame/timecode conversions (library, not wired to
                           a command yet — see its docstring)
  xlsx_utils.py            shared openpyxl helpers (styling, sheet copying)
tests/                      pytest suite + a synthetic fixture generator
data/in/, data/out/          working directories, gitignored
```

## Adding a new operation

Add a function to a new or existing module under `bigmedia/`, wire a
subcommand for it in `cli.py`, and add tests. Keep the same shape as
`sort_workbook()` / `dedupe_workbook()`: take an input path and an output
path, return a small dict of counts/results the CLI can print.
