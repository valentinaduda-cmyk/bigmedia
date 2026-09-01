from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

APP_DIR = Path(__file__).parent

app = FastAPI(title="bigmedia")
app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")


@app.get("/healthz")
def healthz():
    return {"status": "ok"}
