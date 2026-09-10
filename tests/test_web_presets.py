import pytest

from web.presets import (
    PresetExistsError,
    delete_preset,
    get_preset,
    init_db,
    list_presets,
    save_preset,
)


@pytest.fixture
def db_path(tmp_path):
    path = tmp_path / "presets.db"
    init_db(path)
    return path


def test_save_and_get_preset(db_path):
    save_preset(db_path, "sort", "Episode defaults", {"name_column": "Clip Name", "autofit": True})
    result = get_preset(db_path, "sort", "Episode defaults")
    assert result == {"name_column": "Clip Name", "autofit": True}


def test_list_presets_scoped_by_command(db_path):
    save_preset(db_path, "sort", "A", {"name_column": "x"})
    save_preset(db_path, "dedupe", "B", {"name_column": "y"})
    sort_presets = list_presets(db_path, "sort")
    assert [p["name"] for p in sort_presets] == ["A"]


def test_save_duplicate_name_without_overwrite_raises(db_path):
    save_preset(db_path, "sort", "A", {"name_column": "x"})
    with pytest.raises(PresetExistsError):
        save_preset(db_path, "sort", "A", {"name_column": "z"})


def test_save_duplicate_name_with_overwrite_replaces(db_path):
    save_preset(db_path, "sort", "A", {"name_column": "x"})
    save_preset(db_path, "sort", "A", {"name_column": "z"}, overwrite=True)
    assert get_preset(db_path, "sort", "A") == {"name_column": "z"}


def test_delete_preset(db_path):
    save_preset(db_path, "sort", "A", {"name_column": "x"})
    delete_preset(db_path, "sort", "A")
    assert get_preset(db_path, "sort", "A") is None


def test_preset_with_removed_key_loads_without_error(db_path):
    # a preset saved before the output-formatting fields were removed still
    # loads -- unknown keys are simply ignored by _build_kwargs / parse_field.
    save_preset(db_path, "sort", "old", {"name_column": "Clip Name", "autofit": True,
                                         "min_width": 8, "max_width": 60})
    assert get_preset(db_path, "sort", "old")["name_column"] == "Clip Name"
