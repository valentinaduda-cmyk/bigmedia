from web.commands import COMMANDS, SORT_CATEGORIES


def test_all_seven_commands_registered():
    assert set(COMMANDS.keys()) == {
        "sort", "dedupe", "group", "fu-grid", "compare", "fix-getty", "getty-ids",
    }


def test_sort_fields_match_cli_options():
    field_names = {f.name for f in COMMANDS["sort"].fields}
    assert field_names == {
        "name_column", "categories", "skip_categories",
        "autofit", "min_width", "max_width", "uniform_font",
    }


def test_getty_ids_project_name_is_required():
    field = next(f for f in COMMANDS["getty-ids"].fields if f.name == "project_name")
    assert field.required is True


def test_upload_modes():
    assert COMMANDS["sort"].upload_mode == "batch"
    assert COMMANDS["dedupe"].upload_mode == "batch"
    assert COMMANDS["group"].upload_mode == "batch"
    assert COMMANDS["fu-grid"].upload_mode == "batch"
    assert COMMANDS["compare"].upload_mode == "pair"
    assert COMMANDS["fix-getty"].upload_mode == "fix_getty"
    assert COMMANDS["getty-ids"].upload_mode == "combine"


def test_sort_categories_excludes_always_on_fallback():
    assert "3rd parties" not in SORT_CATEGORIES
    assert "AP" in SORT_CATEGORIES


from bigmedia.sort_workbook import analyze_sort


def test_name_and_duration_columns_are_header_dropdowns():
    for slug in ("sort", "group", "fu-grid"):
        by_name = {f.name: f for f in COMMANDS[slug].fields}
        assert by_name["name_column"].options_source == "headers"
    assert {f.name: f for f in COMMANDS["group"].fields}["duration_column"].options_source == "headers"
    # getty-ids and compare share _NAME_COLUMN too — they get the header dropdown for free
    assert {f.name: f for f in COMMANDS["getty-ids"].fields}["name_column"].options_source == "headers"


def test_sheet_fields_are_sheet_dropdowns():
    assert {f.name: f for f in COMMANDS["fu-grid"].fields}["sheet"].options_source == "sheets"
    # group's multi-value "sheets" field stays a text field until group's analyzer lands
    assert {f.name: f for f in COMMANDS["group"].fields}["sheets"].options_source is None


def test_only_sort_has_an_analyzer_for_now():
    assert COMMANDS["sort"].analyze is analyze_sort
    for slug in ("dedupe", "group", "fu-grid", "compare", "fix-getty", "getty-ids"):
        assert COMMANDS[slug].analyze is None


def test_fields_without_options_source_default_to_none():
    assert {f.name: f for f in COMMANDS["sort"].fields}["autofit"].options_source is None
