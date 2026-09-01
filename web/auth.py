import hashlib
import os
import secrets

from fastapi import Depends, Request
from fastapi.responses import RedirectResponse
from starlette.exceptions import HTTPException


def hash_password(plain: str) -> str:
    return hashlib.sha256(plain.encode("utf-8")).hexdigest()


def check_password(plain: str) -> bool:
    expected = os.environ.get("BIGMEDIA_WEB_PASSWORD_HASH", "")
    if not expected:
        return False
    return secrets.compare_digest(hash_password(plain), expected)


class RedirectToLogin(HTTPException):
    def __init__(self):
        super().__init__(status_code=303)


def require_login(request: Request):
    if not request.session.get("authenticated"):
        raise RedirectToLogin()
