import logging
import os
import shutil
import tempfile
from pathlib import Path
from typing import List

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from web.auth import RedirectToLogin, check_password, require_login
from web.commands import COMMANDS
from web.files import parse_field, reject_non_xlsx, zip_files

APP_DIR = Path(__file__).parent

app = FastAPI(title="bigmedia")
app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")

session_secret = os.environ.get("BIGMEDIA_SESSION_SECRET", "dev-only-insecure-secret")
if session_secret == "dev-only-insecure-secret":
    logging.getLogger(__name__).warning(
        "Using default insecure session secret; set BIGMEDIA_SESSION_SECRET env var for production"
    )
app.add_middleware(
    SessionMiddleware,
    secret_key=session_secret,
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


def _command_or_404(slug: str):
    spec = COMMANDS.get(slug)
    if spec is None:
        raise HTTPException(status_code=404)
    return spec


@app.get("/commands/{slug}", dependencies=[Depends(require_login)])
def command_form(request: Request, slug: str):
    spec = _command_or_404(slug)
    return templates.TemplateResponse(
        request, "command.html", {"spec": spec, "commands": COMMANDS, "error": None}
    )


def _save_uploads(files: List[UploadFile], dest_dir: Path) -> list:
    saved = []
    for upload in files:
        path = dest_dir / Path(upload.filename).name
        with open(path, "wb") as out:
            shutil.copyfileobj(upload.file, out)
        saved.append(path)
    return saved


@app.post("/commands/{slug}", dependencies=[Depends(require_login)])
async def command_submit(request: Request, slug: str, files: List[UploadFile] = File(...)):
    spec = _command_or_404(slug)
    bad = reject_non_xlsx([f.filename for f in files])
    form = await request.form()
    if bad:
        return templates.TemplateResponse(
            request, "command.html",
            {"spec": spec, "commands": COMMANDS,
             "error": f"Not an .xlsx file: {', '.join(bad)}"},
        )

    kwargs = {f.name: parse_field(f, form.get(f.name)) for f in spec.fields}
    missing_required = [f.label for f in spec.fields if f.required and not kwargs.get(f.name)]
    if missing_required:
        return templates.TemplateResponse(
            request, "command.html",
            {"spec": spec, "commands": COMMANDS,
             "error": f"Required field(s) missing: {', '.join(missing_required)}"},
        )

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)

        if spec.upload_mode == "batch":
            return _run_batch(spec, files, tmp_dir, kwargs)
        if spec.upload_mode == "combine":
            return _run_combine(spec, files, tmp_dir, kwargs)
        raise HTTPException(status_code=400, detail=f"{slug} not wired up yet")


def _run_batch(spec, files, tmp_dir, kwargs):
    inputs = _save_uploads(files, tmp_dir)
    outputs = []
    for input_path in inputs:
        out_path = tmp_dir / f"{input_path.stem}_{spec.output_suffix}{input_path.suffix}"
        spec.func(str(input_path), str(out_path), **kwargs)
        outputs.append(out_path)

    if len(outputs) == 1:
        return _xlsx_response(outputs[0])
    zip_path = tmp_dir / f"{spec.slug}_results.zip"
    zip_files(outputs, zip_path)
    return _zip_response(zip_path)


def _run_combine(spec, files, tmp_dir, kwargs):
    input_dir = tmp_dir / "input"
    input_dir.mkdir()
    _save_uploads(files, input_dir)
    out_path = tmp_dir / f"{spec.output_suffix}.xlsx"
    project_name = kwargs.pop("project_name")
    spec.func(str(input_dir), str(out_path), project_name, **kwargs)
    return _xlsx_response(out_path)


def _xlsx_response(path: Path):
    return Response(
        content=path.read_bytes(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{path.name}"'},
    )


def _zip_response(path: Path):
    return Response(
        content=path.read_bytes(),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{path.name}"'},
    )
