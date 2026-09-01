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


def _workbook_bytes(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "AP"
    ws.append(["Clip Name"])
    for row in rows:
        ws.append([row])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_compare_form_renders(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/compare/pair")
    assert response.status_code == 200
    assert b"old_file" in response.content
    assert b"new_file" in response.content


def test_compare_runs_and_returns_report(monkeypatch):
    client = _logged_in_client(monkeypatch)
    old_bytes = _workbook_bytes(["clip_a.mov"])
    new_bytes = _workbook_bytes(["clip_a.mov", "clip_b.mov"])
    response = client.post(
        "/commands/compare/pair",
        data={"name_column": "Clip Name"},
        files={
            "old_file": ("old.xlsx", old_bytes, "application/octet-stream"),
            "new_file": ("new.xlsx", new_bytes, "application/octet-stream"),
        },
    )
    assert response.status_code == 200
    assert response.headers["content-disposition"].endswith('"new_vs_old.xlsx"')


def test_fix_getty_form_renders(monkeypatch):
    client = _logged_in_client(monkeypatch)
    response = client.get("/commands/fix-getty/pair")
    assert response.status_code == 200
    assert b"old_file" in response.content
