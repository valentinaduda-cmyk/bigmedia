"""
Sort a media master list into per-source sheets, plus a backup.

Output layout:
  - "Worksheet"        copy of the input, unchanged
  - one sheet per selected category, in CATEGORY_ORDER order, e.g. "AP",
    "Getty Videos", ...

Not every project needs every category (FOX, Veritone and friends are dead
weight on most episodes). categories=/skip_categories= pick which sheets get
written; classification itself is untouched, so a turned-off category's clips
fall into the "3rd parties" fallback sheet for manual review rather than being
dropped. "3rd parties" is that sink and so can never be turned off.

The name column defaults to "Clip Name" — pass name_column= to override for
a workbook that uses a different header. Any "Source Reel Name" column is
left as-is (or absent, if the source doesn't have one) — sort no longer
touches it.
"""
import copy
from openpyxl import load_workbook, Workbook

from .classify import classify, CATEGORY_ORDER, DEFAULT
from .xlsx_utils import (
    find_column_any,
    capture_header_template,
    write_header_row,
    write_data_row,
    copy_sheet_verbatim,
    measure_column_widths,
    apply_column_widths,
    apply_uniform_data_font,
    AUTOFIT_MIN_WIDTH,
    AUTOFIT_MAX_WIDTH,
)

# Real delivery files aren't consistent about the clip-name header: some
# episodes use "Name", others "Clip Name" for the same data. Whichever one
# is passed as name_column is tried first; this covers the other case.
NAME_COLUMN_ALIASES = ("Clip Name", "Name")


def resolve_categories(category_order, categories=None, skip_categories=None):
    """Return the subset of category_order whose sheets should be written.

    Names are matched case-insensitively so callers can pass what a user
    typed; an unrecognized name is an error rather than a silent no-op,
    since a typo would otherwise look like a working filter.
    """
    if categories and skip_categories:
        raise ValueError("Pass categories or skip_categories, not both")

    by_lower = {c.lower(): c for c in category_order}

    def canonical(names):
        resolved = []
        for name in names:
            key = str(name).strip().lower()
            if key not in by_lower:
                raise ValueError(
                    f"Unknown category {name!r}. Known categories: "
                    + ", ".join(category_order)
                )
            resolved.append(by_lower[key])
        return resolved

    if skip_categories:
        skipped = canonical(skip_categories)
        if DEFAULT in skipped:
            raise ValueError(
                f"{DEFAULT!r} cannot be skipped — it collects the clips of "
                "every category that is"
            )
        return [c for c in category_order if c not in skipped]

    if categories:
        kept = set(canonical(categories))
        kept.add(DEFAULT)
        return [c for c in category_order if c in kept]

    return list(category_order)


def count_categories(paths, name_column="Clip Name"):
    """Classify every clip across the given workbooks without writing any
    output, and return {category: count} for every entry in CATEGORY_ORDER
    (0 where a file has none). Used to preview what a real `sort_workbook`
    run would produce, e.g. so a caller can skip presenting empty
    categories as choices."""
    counts = {category: 0 for category in CATEGORY_ORDER}
    for path in paths:
        wb_src = load_workbook(path, read_only=True, data_only=True)
        ws_src = wb_src.active
        name_candidates = [name_column] + [a for a in NAME_COLUMN_ALIASES if a != name_column]
        name_col_idx = find_column_any(ws_src, name_candidates)
        for row in ws_src.iter_rows(min_row=2, min_col=name_col_idx, max_col=name_col_idx):
            counts[classify(row[0].value)] += 1
        wb_src.close()
    return counts


def sort_workbook(src_path, out_path, name_column="Clip Name", category_order=None,
                  categories=None, skip_categories=None,
                  autofit=False, min_width=AUTOFIT_MIN_WIDTH, max_width=AUTOFIT_MAX_WIDTH,
                  uniform_font=False):
    category_order = category_order or CATEGORY_ORDER
    category_order = resolve_categories(category_order, categories, skip_categories)

    wb_src = load_workbook(src_path)
    ws_src = wb_src.active
    max_col = ws_src.max_column
    max_row = ws_src.max_row

    template = capture_header_template(ws_src)
    header_cells = template["header_cells"]
    col_font = template["col_font"]
    col_numfmt = template["col_numfmt"]
    col_widths = template["col_widths"]

    fill_even = copy.copy(ws_src.cell(row=2, column=2).fill)
    fill_odd = copy.copy(ws_src.cell(row=3, column=2).fill)

    name_candidates = [name_column] + [a for a in NAME_COLUMN_ALIASES if a != name_column]
    name_col_idx = find_column_any(ws_src, name_candidates)

    rows_by_cat = {c: [] for c in category_order}
    for r in range(2, max_row + 1):
        name = ws_src.cell(row=r, column=name_col_idx).value
        if name is None or str(name).strip() == "":
            continue
        cat = classify(name)
        if cat not in rows_by_cat:
            # Category turned off for this project: its clips go to the
            # fallback sheet instead of getting one of their own.
            cat = DEFAULT
        rows_by_cat[cat].append(r)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    # "Worksheet" backup tab always goes first.
    ws_backup = wb_out.create_sheet(title="Worksheet")
    copy_sheet_verbatim(ws_src, ws_backup, max_row, max_col)

    # Autofit is measured once, over the whole input, so every category
    # sheet ends up with identical widths -- a workbook whose columns jump
    # around as you switch tabs is worse than one that's uniformly narrow.
    autofit_widths = None
    if autofit:
        autofit_widths = measure_column_widths(
            ([ws_src.cell(row=r, column=c).value for c in range(1, max_col + 1)]
             for r in range(2, max_row + 1)),
            headers=[h.value for h in header_cells],
            min_width=min_width,
            max_width=max_width,
        )

    for cat in category_order:
        ws_out = wb_out.create_sheet(title=cat[:31])
        write_header_row(ws_out, header_cells, ws_src.row_dimensions[1].height)

        for out_r, src_r in enumerate(rows_by_cat.get(cat, []), start=2):
            fill = fill_even if out_r % 2 == 0 else fill_odd
            values = [ws_src.cell(row=src_r, column=c).value for c in range(1, max_col + 1)]
            write_data_row(ws_out, out_r, values, col_font, col_numfmt, fill)

        if autofit_widths:
            apply_column_widths(ws_out, autofit_widths)
        else:
            for col_letter, width in col_widths.items():
                ws_out.column_dimensions[col_letter].width = width
        if uniform_font:
            apply_uniform_data_font(ws_out, col_font[name_col_idx - 1])
        ws_out.freeze_panes = "A2"

    wb_out.save(out_path)
    return {cat: len(rows) for cat, rows in rows_by_cat.items()}
