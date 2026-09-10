import logging
import os
import shutil
import tempfile
from pathlib import Path
from typing import List
from urllib.parse import quote

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from bigmedia.group_duplicates import group_duplicates_workbook

from web.analyze import run_analysis
from web.auth import RedirectToLogin, check_password, require_login
from web.commands import COMMANDS, SORT_CATEGORIES
from web.files import parse_field, reject_non_xlsx, zip_files
from web.presets import PresetExistsError, delete_preset, get_preset, init_db, list_presets, save_preset

APP_DIR = Path(__file__).parent

app = FastAPI(title="bigmedia")
app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")

DB_PATH = Path(os.environ.get("BIGMEDIA_DATA_DIR", "data")) / "presets.db"
init_db(DB_PATH)

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

logger = logging.getLogger(__name__)


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
def home(request: Request):
    return templates.TemplateResponse(request, "home.html", {"commands": COMMANDS})


def _command_or_404(slug: str):
    spec = COMMANDS.get(slug)
    if spec is None:
        raise HTTPException(status_code=404)
    return spec


def _build_kwargs(form, fields) -> dict:
    """Read one value per field from a submitted form, using getlist() for
    "list"-type fields so multiple same-named inputs (checkboxes) all
    survive instead of only the first/last one."""
    return {
        f.name: parse_field(f, form.getlist(f.name) if f.type == "list" else form.get(f.name))
        for f in fields
    }


def _pair_url(slug: str, preset: str = None) -> str:
    url = f"/commands/{slug}/pair"
    if preset:
        url += f"?preset={preset}"
    return url


def _error_response(request: Request, spec, template_name: str, values: dict, error: str, status_code: int = 200):
    return templates.TemplateResponse(
        request, template_name,
        {"spec": spec, "commands": COMMANDS, "presets": list_presets(DB_PATH, spec.slug),
         "values": values, "error": error, "all_categories": SORT_CATEGORIES},
        status_code=status_code,
    )


@app.get("/commands/{slug}", dependencies=[Depends(require_login)])
def command_form(request: Request, slug: str, preset: str = None):
    spec = _command_or_404(slug)
    if spec.upload_mode in ("pair", "fix_getty"):
        return RedirectResponse(url=_pair_url(slug, preset), status_code=303)
    values = get_preset(DB_PATH, slug, preset) if preset else {}
    return templates.TemplateResponse(
        request, "command.html",
        {"spec": spec, "commands": COMMANDS, "presets": list_presets(DB_PATH, slug),
         "values": values, "error": None, "all_categories": SORT_CATEGORIES},
    )


@app.post("/presets/{slug}", dependencies=[Depends(require_login)])
async def save_preset_route(request: Request, slug: str, preset_name: str = Form(...)):
    spec = _command_or_404(slug)
    form = await request.form()
    overwrite = form.get("overwrite") == "on"
    options = _build_kwargs(form, spec.fields)
    is_pair = spec.upload_mode in ("pair", "fix_getty")
    template_name = "pair_command.html" if is_pair else "command.html"
    redirect_url = f"/commands/{slug}/pair" if is_pair else f"/commands/{slug}"
    try:
        save_preset(DB_PATH, slug, preset_name, options, overwrite=overwrite)
    except PresetExistsError:
        return templates.TemplateResponse(
            request, template_name,
            {"spec": spec, "commands": COMMANDS, "presets": list_presets(DB_PATH, slug),
             "values": options, "all_categories": SORT_CATEGORIES,
             "error": f"Preset '{preset_name}' already exists — check 'overwrite' to replace it."},
        )
    return RedirectResponse(url=redirect_url, status_code=303)


@app.post("/presets/{slug}/{name}/delete", dependencies=[Depends(require_login)])
def delete_preset_route(slug: str, name: str):
    spec = _command_or_404(slug)
    delete_preset(DB_PATH, slug, name)
    redirect_url = f"/commands/{slug}/pair" if spec.upload_mode in ("pair", "fix_getty") else f"/commands/{slug}"
    return RedirectResponse(url=redirect_url, status_code=303)


def _save_uploads(files: List[UploadFile], dest_dir: Path) -> list:
    saved = []
    for upload in files:
        path = dest_dir / Path(upload.filename).name
        with open(path, "wb") as out:
            shutil.copyfileobj(upload.file, out)
        saved.append(path)
    return saved


@app.post("/commands/{slug}/analyze", dependencies=[Depends(require_login)])
async def command_analyze(request: Request, slug: str, files: List[UploadFile] = File(...)):
    spec = _command_or_404(slug)
    if spec.upload_mode in ("pair", "fix_getty"):
        raise HTTPException(status_code=404)
    bad = reject_non_xlsx([f.filename for f in files])
    if bad:
        return JSONResponse({"error": f"Not an .xlsx file: {', '.join(bad)}"}, status_code=400)
    form = await request.form()
    with tempfile.TemporaryDirectory() as tmp:
        paths = _save_uploads(files, Path(tmp))
        try:
            data = run_analysis(spec, [str(p) for p in paths], form)
            # Build the response inside the try so a serialization failure
            # (e.g. a non-JSON-safe value slipping through) still degrades to
            # this route's JSON 400 rather than an unhandled 500.
            return JSONResponse(data)
        except Exception as exc:
            logger.exception("analyze failed for %s", slug)
            return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/commands/{slug}", dependencies=[Depends(require_login)])
async def command_submit(request: Request, slug: str, files: List[UploadFile] = File(...)):
    spec = _command_or_404(slug)
    if spec.upload_mode in ("pair", "fix_getty"):
        raise HTTPException(status_code=404)

    form = await request.form()
    kwargs = _build_kwargs(form, spec.fields)

    # The wizard auto-unchecks empty category boxes; if the user approves a
    # form with every box unchecked, `categories` posts nothing and Run would
    # otherwise keep ALL categories (parse_field [] -> None -> keep-all).
    # `categories_present` is a hidden marker the Sort template emits, so we
    # can tell "user cleared every box" from "API caller omitted the field".
    # Skip everything but the always-on "3rd parties" fallback in that case.
    # `categories` and `skip_categories` are mutually exclusive in
    # sort_workbook, so only set skip_categories when categories is falsy.
    if (
        spec.slug == "sort"
        and form.get("categories_present")
        and not kwargs.get("categories")
        and not kwargs.get("skip_categories")
    ):
        kwargs["skip_categories"] = list(SORT_CATEGORIES)

    bad = reject_non_xlsx([f.filename for f in files])
    if bad:
        return _error_response(request, spec, "command.html", kwargs, f"Not an .xlsx file: {', '.join(bad)}")

    missing_required = [f.label for f in spec.fields if f.required and not kwargs.get(f.name)]
    if missing_required:
        return _error_response(
            request, spec, "command.html", kwargs,
            f"Required field(s) missing: {', '.join(missing_required)}",
        )

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)

        if spec.upload_mode == "batch":
            return _run_batch(request, spec, files, tmp_dir, kwargs)
        if spec.upload_mode == "combine":
            return _run_combine(request, spec, files, tmp_dir, kwargs)
        raise HTTPException(status_code=400, detail=f"{slug} not wired up yet")


def _run_batch(request: Request, spec, files, tmp_dir, kwargs):
    inputs = _save_uploads(files, tmp_dir)
    outputs = []
    for input_path in inputs:
        out_path = tmp_dir / f"{input_path.stem}_{spec.output_suffix}{input_path.suffix}"
        try:
            spec.func(str(input_path), str(out_path), **kwargs)
        except Exception as exc:
            logger.exception("Error running %s on %s", spec.slug, input_path.name)
            return _error_response(request, spec, "command.html", kwargs, f"{input_path.name}: {exc}")
        outputs.append(out_path)

    if len(outputs) == 1:
        return _xlsx_response(outputs[0])
    zip_path = tmp_dir / f"{spec.slug}_results.zip"
    zip_files(outputs, zip_path)
    return _zip_response(zip_path)


def _run_combine(request: Request, spec, files, tmp_dir, kwargs):
    input_dir = tmp_dir / "input"
    input_dir.mkdir()
    _save_uploads(files, input_dir)
    out_path = tmp_dir / f"{spec.output_suffix}.xlsx"
    project_name = kwargs.pop("project_name")
    try:
        spec.func(str(input_dir), str(out_path), project_name, **kwargs)
    except Exception as exc:
        logger.exception("Error running %s", spec.slug)
        kwargs["project_name"] = project_name
        offending = files[0].filename if files else spec.slug
        return _error_response(request, spec, "command.html", kwargs, f"{offending}: {exc}")
    return _xlsx_response(out_path)


@app.get("/commands/{slug}/pair", dependencies=[Depends(require_login)])
def pair_form(request: Request, slug: str, preset: str = None):
    spec = _command_or_404(slug)
    if spec.upload_mode not in ("pair", "fix_getty"):
        raise HTTPException(status_code=404)
    values = get_preset(DB_PATH, slug, preset) if preset else {}
    return templates.TemplateResponse(
        request, "pair_command.html",
        {"spec": spec, "commands": COMMANDS, "presets": list_presets(DB_PATH, slug),
         "values": values, "error": None},
    )


@app.post("/commands/{slug}/pair", dependencies=[Depends(require_login)])
async def pair_submit(request: Request, slug: str, old_file: UploadFile = File(...), new_file: UploadFile = File(...)):
    spec = _command_or_404(slug)
    if spec.upload_mode not in ("pair", "fix_getty"):
        raise HTTPException(status_code=404)
    form = await request.form()
    kwargs = _build_kwargs(form, spec.fields)

    bad = reject_non_xlsx([old_file.filename, new_file.filename])
    if bad:
        return _error_response(request, spec, "pair_command.html", kwargs, f"Not an .xlsx file: {', '.join(bad)}")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        old_dir = tmp_dir / "old"
        new_dir = tmp_dir / "new"
        old_dir.mkdir()
        new_dir.mkdir()
        old_path = _save_uploads([old_file], old_dir)[0]
        new_path = _save_uploads([new_file], new_dir)[0]

        try:
            if spec.slug == "compare":
                out_path = tmp_dir / f"{new_path.stem}_{spec.output_suffix}{new_path.suffix}"
                spec.func(str(old_path), str(new_path), str(out_path), **kwargs)
                return _xlsx_response(out_path)

            # fix-getty: mirrors cli.py:cmd_fix_getty
            no_group = kwargs.get("no_group")
            fixed_path = tmp_dir / f"{new_path.stem}_{spec.output_suffix}{new_path.suffix}"
            spec.func(str(old_path), str(new_path), str(fixed_path), name_column=kwargs["name_column"])
            if not no_group:
                grouped_path = tmp_dir / f"{new_path.stem}_{spec.output_suffix}_grouped{new_path.suffix}"
                group_duplicates_workbook(
                    str(fixed_path), str(grouped_path),
                    name_column=kwargs["name_column"],
                    duration_column=kwargs["duration_column"],
                    fps=kwargs["fps"],
                )
                return _xlsx_response(grouped_path)
            return _xlsx_response(fixed_path)
        except Exception as exc:
            logger.exception("Error running %s on %s / %s", spec.slug, old_path.name, new_path.name)
            return _error_response(
                request, spec, "pair_command.html", kwargs,
                f"{old_path.name} / {new_path.name}: {exc}",
            )


def _content_disposition(filename: str) -> str:
    """Build a Content-Disposition header value safe for non-Latin-1 filenames (RFC 5987)."""
    ascii_fallback = filename.encode("ascii", "ignore").decode("ascii") or "download.xlsx"
    return f"attachment; filename*=UTF-8''{quote(filename)}; filename=\"{ascii_fallback}\""


def _xlsx_response(path: Path):
    return Response(
        content=path.read_bytes(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": _content_disposition(path.name)},
    )


def _zip_response(path: Path):
    return Response(
        content=path.read_bytes(),
        media_type="application/zip",
        headers={"Content-Disposition": _content_disposition(path.name)},
    )
