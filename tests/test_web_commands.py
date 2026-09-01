from web.commands import COMMANDS


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
