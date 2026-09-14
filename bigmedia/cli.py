"""
Single entrypoint for every bigmedia operation.

Usage:
    bigmedia sort   <input.xlsx|folder> [-o output] [--name-column "Clip Name"]
                    [--categories "AP,Getty Videos" | --skip-categories "FOX,Veritone"]
    bigmedia dedupe <input.xlsx|folder> [-o output] [--name-column "Clip Name"]
    bigmedia group  <sorted.xlsx|folder> [-o output] [--fps 25]
    bigmedia getty-ids <folder|sorted.xlsx> --project-name "Name" [-o output.xlsx]
    bigmedia fu-grid <sorted.xlsx|folder> [-o output] [--sheet "3rd parties"] [--fps 25]
    bigmedia compare <old.xlsx|folder> <new.xlsx|folder> [-o output]
    bigmedia fix-getty <old folder|file> <new sorted folder|file> [-o output folder] [--no-group]

<input> may be a single .xlsx file or a folder containing several — every
.xlsx file in the folder is processed and each gets its own output file
named after it (e.g. "EP21_master.xlsx" -> "EP21_master_sorted.xlsx").

"sort" writes every category sheet by default. Projects that never use
some sources can trim the output with --skip-categories (drop these sheets)
or --categories (keep only these). Sorting rules are unchanged either way:
a turned-off category's clips land on "3rd parties" for manual review, so
nothing goes missing. "3rd parties" is that sink and can't be turned off.

"group" is a second pass over an already-sorted workbook (the output of
"sort"): on the AP/Getty Videos/Getty Stills/Reuters/Shutterstock/BBC
sheets ("Getty pics" is recognized as an alias for Getty Stills
automatically) it groups duplicate clips together, adds Total
Duration/Seconds columns, and appends a totals summary. Other sheets pass
through untouched.

"getty-ids" reads the "Getty Videos" and "Getty Stills" sheets of every
input workbook (matched case-insensitively; "Getty pics" is recognized as
an alias for the stills sheet automatically, customizable via
--sheet/--stills-sheet for other naming), keeps only unique video clips
with total duration >= --min-seconds (stills are never duration-filtered),
extracts each clip's Getty id from its filename, and writes a report with
videos and stills on separate sheets (--project-name is required and is
also used in the default output filename, "<project name> - Getty_IDs.xlsx").
A clip type's sheet is omitted entirely if it has no clips at all.
--videos-only/--stills-only restrict extraction to one sheet (default: both).

"fu-grid" builds the legal follow-up grid from a sorted workbook's "3rd
parties" sheet: one row per clip (matched by exact clip name), with how
many times it's used and the summed duration of those uses, plus empty
columns the legal team fills in by hand. The source sheet is copied
through untouched as the output's first sheet.

"compare" diffs two versions of the same episode category by category:
how many clips each category holds on either side, which clips the new
version added, which it lost, and which merely changed category (reported
separately, so recategorized clips aren't chased as new material). Point
it at two folders and every episode present in both is compared, paired by
episode number.

"fix-getty" repairs the Getty Videos/Getty Stills split of freshly sorted
workbooks against a previous, hand-checked version of the same episodes
(paired by episode number, not by filename). Stills that came through the
EDL without a .jpg extension land on the videos sheet; this moves them
back. Nothing else in the workbook is touched, and clips the old version
doesn't have are left where they are and listed as unmatched. The result
is re-grouped by default, so the output is a drop-in replacement for a
"group" output.

Add new operations as their own subcommand here as they come up, rather
than as separate one-off scripts — that's the whole point of this file.
"""
import argparse
import sys
from pathlib import Path

from .sort_workbook import sort_workbook
from .dedupe import dedupe_workbook
from .group_duplicates import group_duplicates_workbook
from .getty_ids import build_getty_id_report, getty_ids_filename
from .fu_grid import fu_grid_workbook
from .compare_versions import compare_workbooks
from .getty_split import fix_getty_split, pair_by_episode
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


def _split_list(value):
    if not value:
        return None
    return [part.strip() for part in value.split(",") if part.strip()]


def cmd_sort(args):
    categories = _split_list(args.categories)
    skip_categories = _split_list(args.skip_categories)

    def process_one(input_path, out_path):
        counts = sort_workbook(str(input_path), out_path, name_column=args.name_column,
                               categories=categories, skip_categories=skip_categories)
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


def cmd_fu_grid(args):
    def process_one(input_path, out_path):
        result = fu_grid_workbook(str(input_path), out_path, sheet=args.sheet, name_column=args.name_column, duration_column=args.duration_column, fps=args.fps)
        print(f"Wrote {out_path}")
        print(f"  rows read: {result['rows_in']}")
        print(f"  unique clips: {result['unique_clips']} ({result['multi_use_clips']} used more than once)")

    _run_batch(args, "fu_grid", process_one)


def _print_compare_report(report):
    for title, counts in report.items():
        print(f"  {title}: {counts['old_clips']} clips old -> {counts['new_clips']} clips new"
              f"  (+{len(counts['added'])} added, -{len(counts['removed'])} removed,"
              f" {len(counts['moved_in'])} moved in, {len(counts['moved_out'])} moved out)")
        for label, names in [("added", counts["added"]), ("removed", counts["removed"]),
                             ("moved in", counts["moved_in"]), ("moved out", counts["moved_out"])]:
            for name in names:
                print(f"       {label}: {name}")


def cmd_compare(args):
    old_arg, new_arg = Path(args.old), Path(args.new)
    if old_arg.is_dir() or new_arg.is_dir():
        pairs, unpaired_old, unpaired_new = pair_by_episode(
            _iter_xlsx_inputs(old_arg), _iter_xlsx_inputs(new_arg, exclude_suffix="vs_old"))
        for path in unpaired_new:
            print(f"No old version found for {path.name} — skipped")
        for path in unpaired_old:
            print(f"No new version found for {path.name} — skipped")
        out_dir = Path(args.output) if args.output else (new_arg if new_arg.is_dir() else new_arg.parent) / "compared"
        out_dir.mkdir(parents=True, exist_ok=True)
        for episode, old_path, new_path in pairs:
            out_path = out_dir / f"{new_path.stem}_vs_old.xlsx"
            report = compare_workbooks(str(old_path), str(new_path), str(out_path),
                                       name_column=args.name_column,
                                       case_sensitive=args.case_sensitive, unique=args.unique)
            print(f"Episode {episode}: {new_path.name}  (vs {old_path.name})")
            print(f"  -> {out_path}")
            _print_compare_report(report)
        return

    out_path = args.output or _default_output(args.new, "vs_old")
    report = compare_workbooks(args.old, args.new, out_path, name_column=args.name_column,
                               case_sensitive=args.case_sensitive, unique=args.unique)
    print(f"Wrote {out_path}")
    _print_compare_report(report)


def cmd_fix_getty(args):
    old_paths = _iter_xlsx_inputs(Path(args.old))
    new_paths = _iter_xlsx_inputs(Path(args.new), exclude_suffix="getty_fixed")
    pairs, unpaired_old, unpaired_new = pair_by_episode(old_paths, new_paths)
    for path in unpaired_new:
        print(f"No old version found for {path.name} — skipped")
    for path in unpaired_old:
        print(f"No new version found for {path.name} — skipped")

    out_dir = Path(args.output) if args.output else Path(args.new if Path(args.new).is_dir() else Path(args.new).parent) / "getty_fixed"
    out_dir.mkdir(parents=True, exist_ok=True)

    for episode, old_path, new_path in pairs:
        out_path = out_dir / f"{new_path.stem}_getty_fixed{new_path.suffix}"
        fixed_path = out_path if args.no_group else out_dir / f".{new_path.stem}_getty_fixed_ungrouped{new_path.suffix}"
        result = fix_getty_split(str(old_path), str(new_path), str(fixed_path), name_column=args.name_column)
        if not args.no_group:
            group_duplicates_workbook(str(fixed_path), str(out_path), name_column=args.name_column,
                                      duration_column=args.duration_column, fps=args.fps)
            Path(fixed_path).unlink()

        print(f"Episode {episode}: {new_path.name}  (vs {old_path.name})")
        print(f"  -> {out_path}")
        print(f"  moved to Getty Stills: {len(result['video_to_stills'])} clips / {result['rows_moved_to_stills']} rows"
              + ("  [Getty Stills sheet created]" if result["stills_sheet_created"] else ""))
        for name in result["video_to_stills"]:
            print(f"       {name}")
        print(f"  moved to Getty Videos: {len(result['stills_to_video'])} clips / {result['rows_moved_to_video']} rows")
        for name in result["stills_to_video"]:
            print(f"       {name}")
        unmatched = result["unmatched_videos"] + result["unmatched_stills"]
        print(f"  not on either Getty sheet of the old version, left as sorted: {len(unmatched)} clips")
        for name in unmatched:
            print(f"       {name}")


def cmd_getty_ids(args):
    input_path = Path(args.input)
    if args.output:
        out_path = args.output
    else:
        out_dir = input_path if input_path.is_dir() else input_path.parent
        out_path = str(out_dir / getty_ids_filename(args.project_name))

    counts = build_getty_id_report(
        str(input_path), out_path, args.project_name,
        sheet_name=args.sheet, stills_sheet_name=args.stills_sheet,
        name_column=args.name_column, seconds_column=args.seconds_column,
        min_seconds=args.min_seconds, max_seconds=args.max_seconds,
        include_video=not args.stills_only, include_stills=not args.videos_only,
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
    p_sort.add_argument("--categories", help='Comma-separated categories to keep, e.g. --categories "AP,Getty Videos,Getty Stills". Clips of every other category fall into "3rd parties" (always written). Matched case-insensitively')
    p_sort.add_argument("--skip-categories", help='Comma-separated categories to leave out, e.g. --skip-categories "FOX,Veritone,Content Mint". Their clips fall into "3rd parties". Mutually exclusive with --categories')
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

    p_fu = subparsers.add_parser("fu-grid", help="Build a legal follow-up grid from a sorted workbook's '3rd parties' sheet")
    p_fu.add_argument("input", help="Path to a sorted .xlsx (output of 'sort'), or a folder of them")
    p_fu.add_argument("-o", "--output", help="Output file (single input) or output folder (folder input); default: alongside each input as <name>_fu_grid.xlsx")
    p_fu.add_argument("--sheet", default="3rd parties", help='Sheet to build the grid from, matched case-insensitively (default: "3rd parties")')
    p_fu.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_fu.add_argument("--duration-column", default="Clip Duration", help='Header of the per-row duration column summed per clip (default: "Clip Duration")')
    p_fu.add_argument("--fps", type=int, default=25, help="Frame rate used for timecode math (default: 25)")
    p_fu.set_defaults(func=cmd_fu_grid)

    p_cmp = subparsers.add_parser("compare", help="Per-category diff of two versions of an episode: counts, clips added, clips removed")
    p_cmp.add_argument("old", help="Previous version's .xlsx, or a folder of them")
    p_cmp.add_argument("new", help="New version's .xlsx, or a folder of them")
    p_cmp.add_argument("-o", "--output", help="Output file (single input) or output folder (folder input); default: <new name>_vs_old.xlsx / <new folder>/compared")
    p_cmp.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name", "Name" accepted as an alias)')
    p_cmp.add_argument("--case-sensitive", action="store_true", help="Compare filenames case-sensitively (default: case-insensitive, since old masters uppercase names the new EDL doesn't)")
    p_cmp.add_argument("--unique", action="store_true", help="List each added/removed clip once instead of one row per use in the cut")
    p_cmp.set_defaults(func=cmd_compare)

    p_fix = subparsers.add_parser("fix-getty", help="Re-file Getty Stills wrongly sorted as Getty Videos (and vice versa) using an older, hand-checked version")
    p_fix.add_argument("old", help="Old, hand-checked .xlsx or a folder of them (the truth)")
    p_fix.add_argument("new", help="New sorted .xlsx (output of 'sort') or a folder of them")
    p_fix.add_argument("-o", "--output", help="Output folder (default: <new>/getty_fixed)")
    p_fix.add_argument("--no-group", action="store_true", help="Write the fixed sorted workbook only, skip re-running 'group' on it")
    p_fix.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name", "Name" accepted as an alias)')
    p_fix.add_argument("--duration-column", default="Clip Duration", help='Duration header used when re-grouping (default: "Clip Duration")')
    p_fix.add_argument("--fps", type=int, default=25, help="Frame rate used when re-grouping (default: 25)")
    p_fix.set_defaults(func=cmd_fix_getty)

    p_getty = subparsers.add_parser("getty-ids", help="Extract Getty clip ids from several sorted workbooks into a Getty IDs report")
    p_getty.add_argument("input", help="Path to a folder of sorted .xlsx files (or a single file)")
    p_getty.add_argument("-o", "--output", help='Output report path (default: <input folder>/"<project name> - Getty_IDs.xlsx")')
    p_getty.add_argument("--project-name", required=True, help="Project Name for the report's header and default filename (required)")
    p_getty.add_argument("--sheet", default="Getty Videos", help='Sheet to read video clip names from, matched case-insensitively (default: "Getty Videos")')
    p_getty.add_argument("--stills-sheet", default="Getty Stills", help='Sheet to read stills clip names from, matched case-insensitively (default: "Getty Stills")')
    p_getty.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_getty.add_argument("--seconds-column", default="Seconds", help='Header of the total-duration-in-seconds column (default: "Seconds")')
    p_getty.add_argument("--min-seconds", type=float, default=5, help="Only include VIDEO clips whose total duration is at least this many seconds (default: 5). Stills are never filtered by duration.")
    p_getty.add_argument("--max-seconds", type=float, default=None, help="Only include VIDEO clips whose total duration is below this many seconds (default: no upper bound). Combine with --min-seconds 0 to get clips strictly under a threshold, e.g. --min-seconds 0 --max-seconds 5 for clips under 5 seconds. Stills are never filtered by duration.")
    p_getty_only = p_getty.add_mutually_exclusive_group()
    p_getty_only.add_argument("--videos-only", action="store_true", help="Only extract from the video sheet, skip stills entirely")
    p_getty_only.add_argument("--stills-only", action="store_true", help="Only extract from the stills sheet, skip video entirely")
    p_getty.set_defaults(func=cmd_getty_ids)

    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    sys.exit(main())
