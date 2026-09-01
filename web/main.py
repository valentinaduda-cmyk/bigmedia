import os
from pathlib import Path

from fastapi import Depends, FastAPI, Form, Request
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from web.auth import RedirectToLogin, check_password, require_login

APP_DIR = Path(__file__).parent

app = FastAPI(title="bigmedia")
app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")
app.add_middleware(
    SessionMiddleware,
    secret_key=os.environ.get("BIGMEDIA_SESSION_SECRET", "dev-only-insecure-secret"),
)
templates = Jinja2Templates(directory=APP_DIR / "templates")


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.exception_handler(RedirectToLogin)
def _redirect_to_login(request: Request, exc: RedirectToLogin):
    return RedirectResponse(url="/login", status_code=303)


@app.get("/login")
def login_form(request: Request):
    return templates.TemplateResponse(request, "login.html", {"error": None})


@app.post("/login")
def login_submit(request: Request, password: str = Form(...)):
    if check_password(password):
        request.session["authenticated"] = True
        return RedirectResponse(url="/", status_code=303)
    return templates.TemplateResponse(request, "login.html", {"error": "Incorrect password"}, status_code=200)


@app.get("/", dependencies=[Depends(require_login)])
def home():
    return {"status": "ok"}
