from web.commands import COMMANDS, SORT_CATEGORIES


def test_all_seven_commands_registered():
    assert set(COMMANDS.keys()) == {
        "sort", "dedupe", "group", "fu-grid", "compare", "fix-getty", "getty-ids",
    }


def test_sort_fields_match_cli_options():
    field_names = {f.name for f in COMMANDS["sort"].fields}
    assert field_names == {"name_column", "categories", "skip_categories"}


def test_group_fields_match_cli_options():
    field_names = {f.name for f in COMMANDS["group"].fields}
    assert field_names == {"name_column", "duration_column", "fps", "sheets"}


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
    # group's "sheets" field is now a sheet_checklist with options_source="sheets"
    assert {f.name: f for f in COMMANDS["group"].fields}["sheets"].options_source == "sheets"


def test_only_sort_has_an_analyzer_for_now():
    assert COMMANDS["sort"].analyze is analyze_sort
    for slug in ("dedupe", "group", "fu-grid", "compare", "fix-getty", "getty-ids"):
        assert COMMANDS[slug].analyze is None


def test_fields_without_options_source_default_to_none():
    assert {f.name: f for f in COMMANDS["sort"].fields}["categories"].options_source is None


def test_getty_ids_sheet_fields_are_dropdowns_not_text():
    fields = {f.name: f for f in COMMANDS["getty-ids"].fields}
    assert fields["sheet_name"].options_source == "sheets"
    assert fields["sheet_name"].allow_missing_sheet is False
    assert fields["stills_sheet_name"].options_source == "sheets"
    assert fields["stills_sheet_name"].allow_missing_sheet is True


def test_getty_ids_headers_fields_scope_to_both_sheets():
    fields = {f.name: f for f in COMMANDS["getty-ids"].fields}
    assert fields["name_column"].sheet_source == ["sheet_name", "stills_sheet_name"]
    assert fields["seconds_column"].options_source == "headers"
    assert fields["seconds_column"].sheet_source == ["sheet_name", "stills_sheet_name"]


def test_fu_grid_headers_fields_scope_to_sheet_field():
    fields = {f.name: f for f in COMMANDS["fu-grid"].fields}
    assert fields["name_column"].sheet_source == ["sheet"]
    assert fields["duration_column"].sheet_source == ["sheet"]


def test_group_sheets_field_is_a_checklist_scoping_name_and_duration():
    fields = {f.name: f for f in COMMANDS["group"].fields}
    assert fields["sheets"].type == "sheet_checklist"
    assert fields["sheets"].options_source == "sheets"
    assert fields["name_column"].sheet_source == ["sheets"]
    assert fields["duration_column"].sheet_source == ["sheets"]


def test_sort_and_dedupe_fields_declare_no_sheet_source():
    # Explicitly unchanged: no sheet_source anywhere for the two commands
    # this feature does not touch.
    for slug in ("sort", "dedupe"):
        for f in COMMANDS[slug].fields:
            assert not f.sheet_source
