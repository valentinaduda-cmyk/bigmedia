from fastapi.testclient import TestClient

from web.auth import hash_password
from web.main import app


def _client(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    c = TestClient(app)
    c.post("/login", data={"password": "pw"})
    return c


def test_name_column_renders_as_analyze_select(monkeypatch):
    html = _client(monkeypatch).get("/commands/dedupe").text
    assert '<select name="name_column"' in html
    assert 'data-analyze' in html
    assert 'data-source="headers"' in html


def test_form_has_analyze_url_and_script(monkeypatch):
    html = _client(monkeypatch).get("/commands/dedupe").text
    assert 'data-analyze-url="/commands/dedupe/analyze"' in html
    assert '/static/analyze.js' in html


def test_warning_banner_present_and_hidden(monkeypatch):
    html = _client(monkeypatch).get("/commands/sort").text
    assert 'id="analyze-warnings"' in html
    assert 'hidden' in html.split('id="analyze-warnings"')[1][:40]


def test_sort_still_has_category_checkboxes(monkeypatch):
    html = _client(monkeypatch).get("/commands/sort").text
    assert 'name="categories"' in html
    assert 'class="category-name"' in html


def test_fu_grid_sheet_field_is_sheet_select(monkeypatch):
    html = _client(monkeypatch).get("/commands/fu-grid").text
    assert '<select name="sheet"' in html
    assert 'data-source="sheets"' in html


def test_group_sheets_field_is_sheet_select(monkeypatch):
    # group's "sheets" field is now a sheet_checklist with select dropdown
    html = _client(monkeypatch).get("/commands/group").text
    assert '<select name="sheets"' in html
    assert 'data-source="sheets"' in html
