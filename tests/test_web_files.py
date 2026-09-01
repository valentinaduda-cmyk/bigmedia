import zipfile

from web.commands import FieldSpec
from web.files import parse_field, reject_non_xlsx, zip_files


def test_reject_non_xlsx_flags_bad_files():
    bad = reject_non_xlsx(["a.xlsx", "b.csv", "c.XLSX"])
    assert bad == ["b.csv"]


def test_reject_non_xlsx_all_valid_returns_empty():
    assert reject_non_xlsx(["a.xlsx", "b.xlsx"]) == []


def test_parse_field_list_splits_and_strips():
    field = FieldSpec("categories", "Categories", "list")
    assert parse_field(field, "AP, Getty Videos ,  ") == ["AP", "Getty Videos"]


def test_parse_field_list_blank_is_none():
    field = FieldSpec("categories", "Categories", "list")
    assert parse_field(field, "") is None


def test_parse_field_list_accepts_multiple_checkbox_values():
    field = FieldSpec("categories", "Categories", "list")
    assert parse_field(field, ["AP", "GFX"]) == ["AP", "GFX"]


def test_parse_field_list_empty_checkbox_list_is_none():
    field = FieldSpec("categories", "Categories", "list")
    assert parse_field(field, []) is None


def test_parse_field_number_blank_uses_default_none():
    field = FieldSpec("max_seconds", "Max seconds", "number")
    assert parse_field(field, "") is None


def test_parse_field_number_parses_float_and_int():
    field = FieldSpec("fps", "FPS", "number", 25)
    assert parse_field(field, "30") == 30
    assert parse_field(field, "29.97") == 29.97


def test_parse_field_checkbox_present_or_absent():
    field = FieldSpec("autofit", "Autofit", "checkbox", False)
    assert parse_field(field, "on") is True
    assert parse_field(field, None) is False


def test_parse_field_text_passthrough():
    field = FieldSpec("name_column", "Name column", "text", "Clip Name")
    assert parse_field(field, "Filename") == "Filename"


def test_zip_files_creates_archive_with_all_members(tmp_path):
    f1 = tmp_path / "a.xlsx"
    f2 = tmp_path / "b.xlsx"
    f1.write_bytes(b"one")
    f2.write_bytes(b"two")
    zip_path = tmp_path / "out.zip"
    zip_files([f1, f2], zip_path)
    with zipfile.ZipFile(zip_path) as zf:
        assert sorted(zf.namelist()) == ["a.xlsx", "b.xlsx"]
