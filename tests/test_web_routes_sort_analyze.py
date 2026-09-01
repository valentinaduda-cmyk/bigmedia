import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from web.auth import hash_password
from web.main import app


def _logged_in_client(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    client = TestClient(app)
    client.post("/login", data={"password": "pw"})
    return client


def _xlsx_bytes(names):
    wb = Workbook()
    ws = wb.active
    ws.append(["Clip Name"])
    for name in names:
        ws.append([name])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_analyze_returns_category_counts(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _xlsx_bytes(["BM1234_something.mxf", "BM5678_other.mxf", "random_file_xyz.mov"])
    files = {"files": ("master.xlsx", content, "application/octet-stream")}
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 200
    counts = response.json()["counts"]
    assert counts["AP"] == 2
    assert counts["3rd parties"] == 1
    assert counts["Reuters"] == 0


def test_analyze_rejects_non_xlsx(monkeypatch):
    client = _logged_in_client(monkeypatch)
    files = {"files": ("notes.txt", b"hello", "text/plain")}
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 400
    assert "notes.txt" in response.json()["error"]


def test_sort_accepts_multiple_categories_checkboxes(monkeypatch):
    client = _logged_in_client(monkeypatch)
    content = _xlsx_bytes([
        "BM1234_something.mxf",              # AP
        "045_SO_EP18_01_3DExplainer_TXLS.mov",  # GFX
        "shutterstock_777.mp4",              # Shutterstock
    ])
    files = {"files": ("master.xlsx", content, "application/octet-stream")}
    response = client.post(
        "/commands/sort",
        data={"name_column": "Clip Name", "categories": ["AP", "GFX"]},
        files=files,
    )
    assert response.status_code == 200
    wb = load_workbook(io.BytesIO(response.content))
    # Both checked categories are kept, the unchecked one falls to the fallback sheet.
    assert "AP" in wb.sheetnames
    assert "GFX" in wb.sheetnames
    assert "Shutterstock" not in wb.sheetnames
    assert "3rd parties" in wb.sheetnames


def test_analyze_requires_login(monkeypatch):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    client = TestClient(app, follow_redirects=False)
    files = {"files": ("master.xlsx", _xlsx_bytes(["x.mov"]), "application/octet-stream")}
    response = client.post("/commands/sort/analyze", data={"name_column": "Clip Name"}, files=files)
    assert response.status_code == 303
