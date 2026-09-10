from dataclasses import dataclass
from typing import Any, Callable

from bigmedia.classify import CATEGORY_ORDER, DEFAULT
from bigmedia.compare_versions import compare_workbooks
from bigmedia.dedupe import dedupe_workbook
from bigmedia.fu_grid import fu_grid_workbook
from bigmedia.getty_ids import build_getty_id_report
from bigmedia.getty_split import fix_getty_split
from bigmedia.group_duplicates import group_duplicates_workbook
from bigmedia.sort_workbook import sort_workbook, analyze_sort


@dataclass
class FieldSpec:
    name: str
    label: str
    type: str  # "text" | "number" | "checkbox" | "list"
    default: Any = None
    required: bool = False
    options_source: str = None  # None | "headers" | "sheets"


@dataclass
class CommandSpec:
    slug: str
    title: str
    upload_mode: str  # "batch" | "combine" | "pair" | "fix_getty"
    func: Callable
    fields: list
    output_suffix: str
    analyze: Callable = None


_NAME_COLUMN = FieldSpec("name_column", "Filename column", "text", "Clip Name", options_source="headers")
_DURATION_COLUMN = FieldSpec("duration_column", "Duration column", "text", "Clip Duration", options_source="headers")
_FPS = FieldSpec("fps", "Frame rate (fps)", "number", 25)

COMMANDS = {
    "sort": CommandSpec(
        slug="sort", title="Sort", upload_mode="batch", func=sort_workbook,
        output_suffix="sorted", analyze=analyze_sort,
        fields=[
            _NAME_COLUMN,
            FieldSpec("categories", "Categories to keep (comma-separated, blank = all)", "list"),
            FieldSpec("skip_categories", "Categories to skip (comma-separated)", "list"),
        ],
    ),
    "dedupe": CommandSpec(
        slug="dedupe", title="Dedupe", upload_mode="batch", func=dedupe_workbook,
        output_suffix="deduped",
        fields=[_NAME_COLUMN],
    ),
    "group": CommandSpec(
        slug="group", title="Group", upload_mode="batch", func=group_duplicates_workbook,
        output_suffix="grouped",
        fields=[
            _NAME_COLUMN, _DURATION_COLUMN, _FPS,
            FieldSpec("sheets", "Sheets to group (comma-separated, blank = default set)", "list"),
        ],
    ),
    "fu-grid": CommandSpec(
        slug="fu-grid", title="FU Grid", upload_mode="batch", func=fu_grid_workbook,
        output_suffix="fu_grid",
        fields=[
            FieldSpec("sheet", "Source sheet", "text", "3rd parties", options_source="sheets"),
            _NAME_COLUMN, _DURATION_COLUMN, _FPS,
        ],
    ),
    "compare": CommandSpec(
        slug="compare", title="Compare", upload_mode="pair", func=compare_workbooks,
        output_suffix="vs_old",
        fields=[
            _NAME_COLUMN,
            FieldSpec("case_sensitive", "Case-sensitive filename match", "checkbox", False),
            FieldSpec("unique", "List each clip once instead of once per use", "checkbox", False),
        ],
    ),
    "fix-getty": CommandSpec(
        slug="fix-getty", title="Fix Getty split", upload_mode="fix_getty", func=fix_getty_split,
        output_suffix="getty_fixed",
        fields=[
            _NAME_COLUMN, _DURATION_COLUMN, _FPS,
            FieldSpec("no_group", "Skip re-grouping after fix", "checkbox", False),
        ],
    ),
    "getty-ids": CommandSpec(
        slug="getty-ids", title="Getty IDs", upload_mode="combine", func=build_getty_id_report,
        output_suffix="getty_ids",
        fields=[
            FieldSpec("project_name", "Project name", "text", required=True),
            FieldSpec("production_company", "Production company", "text", "KM Record a.s./Big Media"),
            FieldSpec("broadcaster", "Broadcaster", "text", ""),
            FieldSpec("rights", "Rights requested", "text", "in perpetuity/worldwide/all media"),
            FieldSpec("sheet_name", "Video sheet name", "text", "Getty Videos"),
            FieldSpec("stills_sheet_name", "Stills sheet name", "text", "Getty Stills"),
            _NAME_COLUMN,
            FieldSpec("seconds_column", "Seconds column", "text", "Seconds"),
            FieldSpec("min_seconds", "Min seconds (videos only)", "number", 5),
            FieldSpec("max_seconds", "Max seconds (videos only, blank = no limit)", "number"),
            FieldSpec("include_video", "Include video clips", "checkbox", True),
            FieldSpec("include_stills", "Include stills", "checkbox", True),
        ],
    ),
}

# The sort form renders these as checkboxes. "3rd parties" is excluded --
# it's the always-on fallback sheet and can't be turned off (matches the
# CLI's own resolve_categories() rule).
SORT_CATEGORIES = [c for c in CATEGORY_ORDER if c != DEFAULT]
