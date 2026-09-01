import hashlib
import os

from fastapi.testclient import TestClient

from web.auth import hash_password
from web.main import app


def _set_password(monkeypatch, plain):
    monkeypatch.setenv("BIGMEDIA_WEB_PASSWORD_HASH", hash_password(plain))


def test_hash_password_is_sha256_hex():
    assert hash_password("secret") == hashlib.sha256(b"secret").hexdigest()


def test_protected_route_redirects_when_not_logged_in():
    client = TestClient(app, follow_redirects=False)
    response = client.get("/")
    assert response.status_code in (302, 303, 307)
    assert response.headers["location"] == "/login"


def test_login_with_correct_password_grants_access(monkeypatch):
    _set_password(monkeypatch, "correct-horse")
    client = TestClient(app)
    response = client.post("/login", data={"password": "correct-horse"}, follow_redirects=False)
    assert response.status_code in (302, 303)
    home = client.get("/")
    assert home.status_code == 200


def test_login_with_wrong_password_rejected(monkeypatch):
    _set_password(monkeypatch, "correct-horse")
    client = TestClient(app)
    response = client.post("/login", data={"password": "wrong"})
    assert response.status_code == 200
    assert b"Incorrect password" in response.content
