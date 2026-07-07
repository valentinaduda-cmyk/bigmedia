"""
Single entrypoint for every bigmedia operation.

Usage:
    bigmedia sort   <input.xlsx|folder> [-o output] [--name-column "Clip Name"]
    bigmedia dedupe <input.xlsx|folder> [-o output] [--name-column "Clip Name"]
    bigmedia group  <sorted.xlsx|folder> [-o output] [--fps 25]
    bigmedia getty-ids <folder|sorted.xlsx> [-o output.xlsx]

<input> may be a single .xlsx file or a folder containing several — every
.xlsx file in the folder is processed and each gets its own output file
named after it (e.g. "EP21_master.xlsx" -> "EP21_master_sorted.xlsx").

"group" is a second pass over an already-sorted workbook (the output of
"sort"): on the AP/Getty Videos/Getty Stills/Reuters/Shutterstock/BBC
sheets it groups duplicate clips together, adds Total Duration/Seconds
columns, and appends a totals summary. Other sheets pass through untouched.

"getty-ids" reads the "Getty videos" and "Getty pics" sheets of every
input workbook (customizable via --sheet/--stills-sheet), keeps only
unique clips with total duration >= --min-seconds, extracts each clip's
Getty id from its filename, and writes one Customer Declaration Form-
styled report (--project-name is required; --production-company,
--broadcaster, --rights fill the rest of the form's header).

Add new operations as their own subcommand here as they come up, rather
than as separate one-off scripts — that's the whole point of this file.
"""
import argparse
import sys
from pathlib import Path

from .sort_workbook import sort_workbook
from .dedupe import dedupe_workbook
from .group_duplicates import group_duplicates_workbook
from .getty_ids import build_getty_id_report
from .xlsx_utils import iter_xlsx_files as _iter_xlsx_inputs


def _default_output(input_path, suffix: str) -> str:
    p = Path(input_path)
    return str(p.with_name(f"{p.stem}_{suffix}{p.suffix}"))


def _resolve_output(input_path: Path, output_arg, suffix: str, batch: bool) -> str:
    if batch:
        out_dir = Path(output_arg) if output_arg else input_path.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        return str(out_dir / f"{input_path.stem}_{suffix}{input_path.suffix}")
    return output_arg or _default_output(input_path, suffix)


def _run_batch(args, suffix, process_one):
    input_arg = Path(args.input)
    batch = input_arg.is_dir()
    inputs = _iter_xlsx_inputs(input_arg, exclude_suffix=suffix)
    if not inputs:
        print(f"No .xlsx files found in {input_arg}")
        return
    for input_path in inputs:
        out_path = _resolve_output(input_path, args.output, suffix, batch)
        process_one(input_path, out_path)


def cmd_sort(args):
    def process_one(input_path, out_path):
        counts = sort_workbook(str(input_path), out_path, name_column=args.name_column)
        print(f"Wrote {out_path}")
        for cat, n in counts.items():
            print(f"  {cat}: {n}")

    _run_batch(args, "sorted", process_one)


def cmd_dedupe(args):
    def process_one(input_path, out_path):
        result = dedupe_workbook(str(input_path), out_path, name_column=args.name_column)
        print(f"Wrote {out_path}")
        print(f"  kept: {result['kept']}")
        print(f"  dropped as duplicates: {result['dropped']}")

    _run_batch(args, "deduped", process_one)


def cmd_group(args):
    sheets = [s.strip() for s in args.sheets.split(",")] if args.sheets else None

    def process_one(input_path, out_path):
        result = group_duplicates_workbook(str(input_path), out_path, sheets=sheets, name_column=args.name_column, duration_column=args.duration_column, fps=args.fps)
        print(f"Wrote {out_path}")
        for sheet, counts in result.items():
            print(f"  {sheet}: {counts['total_clips']} clips, {counts['total_unique_clips']} unique, "
                  f"{counts['total_clips_over_4s']} >4s, {counts['total_seconds']}s total "
                  f"({counts['total_seconds_over_4s']}s from clips >4s)")

    _run_batch(args, "grouped", process_one)


def cmd_getty_ids(args):
    input_path = Path(args.input)
    if args.output:
        out_path = args.output
    elif input_path.is_dir():
        out_path = str(input_path / "getty_ids.xlsx")
    else:
        out_path = _default_output(input_path, "getty_ids")

    counts = build_getty_id_report(
        str(input_path), out_path, args.project_name,
        sheet_name=args.sheet, stills_sheet_name=args.stills_sheet,
        name_column=args.name_column, seconds_column=args.seconds_column,
        min_seconds=args.min_seconds,
        production_company=args.production_company,
        broadcaster=args.broadcaster, rights=args.rights,
    )
    print(f"Wrote {out_path}")
    for name, n in counts.items():
        print(f"  {name}: {n['video']} video, {n['stills']} stills")


def build_parser():
    parser = argparse.ArgumentParser(prog="bigmedia", description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_sort = subparsers.add_parser("sort", help="Classify clips into per-source sheets")
    p_sort.add_argument("input", help="Path to a master .xlsx, or a folder of .xlsx files")
    p_sort.add_argument("-o", "--output", help="Output file (single input) or output folder (folder input); default: alongside each input as <name>_sorted.xlsx")
    p_sort.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_sort.set_defaults(func=cmd_sort)

    p_dedupe = subparsers.add_parser("dedupe", help="Split a workbook into deduped + duplicate rows")
    p_dedupe.add_argument("input", help="Path to a master .xlsx, or a folder of .xlsx files")
    p_dedupe.add_argument("-o", "--output", help="Output file (single input) or output folder (folder input); default: alongside each input as <name>_deduped.xlsx")
    p_dedupe.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_dedupe.set_defaults(func=cmd_dedupe)

    p_group = subparsers.add_parser("group", help="Group duplicate clips on select sheets of an already-sorted workbook")
    p_group.add_argument("input", help="Path to a sorted .xlsx (output of 'sort'), or a folder of them")
    p_group.add_argument("-o", "--output", help="Output file (single input) or output folder (folder input); default: alongside each input as <name>_grouped.xlsx")
    p_group.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_group.add_argument("--duration-column", default="Clip Duration", help='Header of the per-row duration column summed into "Total Duration" (default: "Clip Duration")')
    p_group.add_argument("--fps", type=int, default=25, help="Frame rate used for timecode math (default: 25)")
    p_group.add_argument("--sheets", help='Comma-separated sheet names to group (default: "AP,Getty Videos,Getty Stills,Reuters,Shutterstock,BBC", matched case-insensitively). Use this when a file names its sheets differently, e.g. --sheets "AP,Getty videos,Getty pics,Reuters,Shutterstock"')
    p_group.set_defaults(func=cmd_group)

    p_getty = subparsers.add_parser("getty-ids", help="Extract Getty clip ids from several sorted workbooks into a Customer Declaration Form report")
    p_getty.add_argument("input", help="Path to a folder of sorted .xlsx files (or a single file)")
    p_getty.add_argument("-o", "--output", help="Output report path (default: <input folder>/getty_ids.xlsx)")
    p_getty.add_argument("--project-name", required=True, help="Project Name for the form's header (required)")
    p_getty.add_argument("--production-company", default="KM Record a.s./Big Media", help='Production Company for the form\'s header (default: "KM Record a.s./Big Media")')
    p_getty.add_argument("--broadcaster", default="", help="Broadcaster for the form's header (default: blank)")
    p_getty.add_argument("--rights", default="in perpetuity/worldwide/all media", help='Rights Requested for the form\'s header (default: "in perpetuity/worldwide/all media")')
    p_getty.add_argument("--sheet", default="Getty videos", help='Sheet to read video clip names from (default: "Getty videos")')
    p_getty.add_argument("--stills-sheet", default="Getty pics", help='Sheet to read stills clip names from (default: "Getty pics")')
    p_getty.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_getty.add_argument("--seconds-column", default="Seconds", help='Header of the total-duration-in-seconds column (default: "Seconds")')
    p_getty.add_argument("--min-seconds", type=float, default=5, help="Only include clips whose total duration is at least this many seconds (default: 5)")
    p_getty.set_defaults(func=cmd_getty_ids)

    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    sys.exit(main())
