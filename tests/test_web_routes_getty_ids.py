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


def _sorted_workbook_bytes():
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Getty Videos")
    ws.append(["Clip Name", "Seconds"])
    ws.append(["GettyImages-12345.mov", 10])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_getty_ids_requires_project_name(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("ep1.xlsx", _sorted_workbook_bytes(), "application/octet-stream")}
    response = client.post("/commands/getty-ids", data={}, files=files)
    assert response.status_code == 200
    assert b"error-banner" in response.content


def test_getty_ids_combines_multiple_files_into_one_report(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _sorted_workbook_bytes()
    files = [
        ("files", ("ep1.xlsx", content, "application/octet-stream")),
        ("files", ("ep2.xlsx", content, "application/octet-stream")),
    ]
    response = client.post(
        "/commands/getty-ids", data={"project_name": "Test Project"}, files=files
    )
    assert response.status_code == 200
    assert response.headers["content-disposition"].endswith('"getty_ids.xlsx"')
