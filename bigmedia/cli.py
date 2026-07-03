"""
Single entrypoint for every bigmedia operation.

Usage:
    bigmedia sort   <input.xlsx> [-o output.xlsx] [--name-column "Clip Name"]
    bigmedia dedupe <input.xlsx> [-o output.xlsx] [--name-column "Clip Name"]

Add new operations as their own subcommand here as they come up, rather
than as separate one-off scripts — that's the whole point of this file.
"""
import argparse
import sys
from pathlib import Path

from .sort_workbook import sort_workbook
from .dedupe import dedupe_workbook


def _default_output(input_path: str, suffix: str) -> str:
    p = Path(input_path)
    return str(p.with_name(f"{p.stem}_{suffix}{p.suffix}"))


def cmd_sort(args):
    out_path = args.output or _default_output(args.input, "sorted")
    counts = sort_workbook(args.input, out_path, name_column=args.name_column)
    print(f"Wrote {out_path}")
    for cat, n in counts.items():
        print(f"  {cat}: {n}")


def cmd_dedupe(args):
    out_path = args.output or _default_output(args.input, "deduped")
    result = dedupe_workbook(args.input, out_path, name_column=args.name_column)
    print(f"Wrote {out_path}")
    print(f"  kept: {result['kept']}")
    print(f"  dropped as duplicates: {result['dropped']}")


def build_parser():
    parser = argparse.ArgumentParser(prog="bigmedia", description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_sort = subparsers.add_parser("sort", help="Classify clips into per-source sheets")
    p_sort.add_argument("input", help="Path to the master .xlsx")
    p_sort.add_argument("-o", "--output", help="Output path (default: <input>_sorted.xlsx)")
    p_sort.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_sort.set_defaults(func=cmd_sort)

    p_dedupe = subparsers.add_parser("dedupe", help="Split a workbook into deduped + duplicate rows")
    p_dedupe.add_argument("input", help="Path to the master .xlsx")
    p_dedupe.add_argument("-o", "--output", help="Output path (default: <input>_deduped.xlsx)")
    p_dedupe.add_argument("--name-column", default="Clip Name", help='Header of the filename column (default: "Clip Name")')
    p_dedupe.set_defaults(func=cmd_dedupe)

    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    sys.exit(main())
