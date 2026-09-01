import io

from fastapi.testclient import TestClient
from openpyxl import Workbook

from web.auth import hash_password
from web.main import app


def _logged_in_client(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    client = TestClient(app)
    client.post("/login", data={"password": "pw"})
    return client


def _sample_xlsx_bytes(column="Clip Name", rows=("random_file_xyz.mov",)):
    wb = Workbook()
    ws = wb.active
    ws.append([column])
    for row in rows:
        ws.append([row])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --- Finding 1: nav / flat-URL guard for pair-mode commands -----------------

def test_get_flat_url_for_compare_redirects_to_pair(monkeypatch):
    client = TestClient(app, follow_redirects=False)
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    client.post("/login", data={"password": "pw"})
    response = client.get("/commands/compare")
    assert response.status_code in (303, 404)
    if response.status_code == 303:
        assert response.headers["location"].endswith("/commands/compare/pair")


def test_post_flat_url_for_fix_getty_is_rejected(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("f.xlsx", _sample_xlsx_bytes(), "application/octet-stream")}
    response = client.post("/commands/fix-getty", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 404


def test_nav_links_to_pair_url_for_compare(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/sort")
    assert response.status_code == 200
    assert b'href="/commands/compare/pair"' in response.content
    assert b'href="/commands/fix-getty/pair"' in response.content
    assert b'href="/commands/compare"' not in response.content


# --- Finding 2: business-function exceptions render a clean error ----------

def test_bad_workbook_content_renders_error_not_500(monkeypatch):
    client = _logged_in_client(monkeypatch)
    # Valid xlsx but missing the configured name column -> business function should raise
    bad_wb = _sample_xlsx_bytes(column="Some Other Column")
    files = {"files": ("bad.xlsx", bad_wb, "application/octet-stream")}
    response = client.post("/commands/sort", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    assert b"error-banner" in response.content
    assert b"bad.xlsx" in response.content


# --- Finding 3: non-Latin-1 filenames don't crash the download -------------

def test_non_ascii_filename_round_trips(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("Přehled.xlsx", _sample_xlsx_bytes(), "application/octet-stream")}
    response = client.post("/commands/sort", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    cd = response.headers["content-disposition"]
    assert "filename*=UTF-8''" in cd


# --- Finding 4: error re-renders preserve user input and presets -----------

def test_bad_upload_error_preserves_field_values(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("notes.txt", b"hello", "text/plain")}
    response = client.post("/commands/sort", data={"name_column": "Custom"}, files=files)
    assert response.status_code == 200
    assert b'value="Custom"' in response.content


# --- Finding 5: home page lists commands instead of JSON stub --------------

def test_home_page_lists_commands(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/")
    assert response.status_code == 200
    assert response.headers["content-type"].startswith("text/html")
    assert b"/commands/sort" in response.content
    assert b"/commands/compare/pair" in response.content
