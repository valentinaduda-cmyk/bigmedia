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


def _sample_xlsx_bytes():
    wb = Workbook()
    ws = wb.active
    ws.append(["Clip Name"])
    ws.append(["random_file_xyz.mov"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_sort_form_renders(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/sort")
    assert response.status_code == 200
    assert b"name_column" in response.content


def test_sort_form_renders_category_checkboxes_all_checked_by_default(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/sort")
    body = response.content.decode()
    assert '<input type="checkbox" name="categories" value="AP" checked>' in body
    assert '<input type="checkbox" name="categories" value="Reuters" checked>' in body
    # The always-on fallback isn't a togglable checkbox.
    assert 'value="3rd parties"' not in body


def test_sort_single_file_returns_xlsx(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("master.xlsx", _sample_xlsx_bytes(), "application/octet-stream")}
    response = client.post("/commands/sort", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    assert response.headers["content-type"].startswith("application/")
    assert response.headers["content-disposition"].endswith('"master_sorted.xlsx"')


def test_sort_multiple_files_returns_zip(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _sample_xlsx_bytes()
    files = [
        ("files", ("ep1.xlsx", content, "application/octet-stream")),
        ("files", ("ep2.xlsx", content, "application/octet-stream")),
    ]
    response = client.post("/commands/sort", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/zip"


def test_sort_rejects_non_xlsx(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("notes.txt", b"hello", "text/plain")}
    response = client.post("/commands/sort", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    assert b"notes.txt" in response.content
    assert b"error-banner" in response.content


def test_unauthenticated_request_redirected(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    client = TestClient(app, follow_redirects=False)
    response = client.get("/commands/sort")
    assert response.status_code == 303


def test_sort_path_traversal_blocked(monkeypatch):
    """Verify that path traversal sequences in filenames are sanitized."""
    client = _logged_in_client(monkeypatch)
    # Try to upload a file with path traversal sequences
    # The filename should be sanitized to just "evil.xlsx"
    files = {"files": ("../evil.xlsx", _sample_xlsx_bytes(), "application/octet-stream")}
    response = client.post("/commands/sort", data={"name_column": "Clip Name"}, files=files)
    # Request should succeed
    assert response.status_code == 200
    # Response should contain the sorted output (not an error)
    assert response.headers["content-disposition"].endswith('"evil_sorted.xlsx"')
