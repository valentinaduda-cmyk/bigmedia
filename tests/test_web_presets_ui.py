from fastapi.testclient import TestClient

from web.auth import hash_password
from web.main import DB_PATH, app
from web.presets import init_db


def _logged_in_client(monkeypatch, tmp_path):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    db_path = tmp_path / "presets.db"
    monkeypatch.setattr("web.main.DB_PATH", db_path)
    init_db(db_path)
    client = TestClient(app)
    client.post("/login", data={"password": "pw"})
    return client


def test_save_preset_then_appears_on_form(monkeypatch, tmp_path):
    client = _logged_in_client(monkeypatch, tmp_path)
    response = client.post(
        "/presets/sort",
        data={"preset_name": "Episode defaults", "name_column": "Clip Name"},
        follow_redirects=False,
    )
    assert response.status_code == 303
    form = client.get("/commands/sort")
    assert b"Episode defaults" in form.content


def test_saving_duplicate_name_shows_overwrite_prompt(monkeypatch, tmp_path):
    client = _logged_in_client(monkeypatch, tmp_path)
    client.post("/presets/sort", data={"preset_name": "A", "name_column": "Clip Name"})
    response = client.post("/presets/sort", data={"preset_name": "A", "name_column": "Other"})
    assert response.status_code == 200
    assert b"already exists" in response.content


def test_delete_preset(monkeypatch, tmp_path):
    client = _logged_in_client(monkeypatch, tmp_path)
    client.post("/presets/sort", data={"preset_name": "A", "name_column": "Clip Name"})
    client.post("/presets/sort/A/delete")
    form = client.get("/commands/sort")
    assert b">A<" not in form.content


def test_loading_preset_prefills_form(monkeypatch, tmp_path):
    client = _logged_in_client(monkeypatch, tmp_path)
    client.post("/presets/sort", data={"preset_name": "A", "name_column": "Custom Col"})
    response = client.get("/commands/sort?preset=A")
    assert b'value="Custom Col"' in response.content


def test_preset_routes_require_login(monkeypatch, tmp_path):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password("pw"))
    db_path = tmp_path / "presets.db"
    monkeypatch.setattr("web.main.DB_PATH", db_path)
    init_db(db_path)
    client = TestClient(app)
    response = client.post(
        "/presets/sort",
        data={"preset_name": "A", "name_column": "Clip Name"},
        follow_redirects=False,
    )
    assert response.status_code == 303
    assert response.headers["location"] == "/login"

    response = client.post("/presets/sort/A/delete", follow_redirects=False)
    assert response.status_code == 303
    assert response.headers["location"] == "/login"
