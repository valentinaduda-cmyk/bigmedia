import json
import sqlite3
from pathlib import Path


class PresetExistsError(Exception):
    pass


def _connect(db_path: Path) -> sqlite3.Connection:
    return sqlite3.connect(db_path)


def init_db(db_path: Path) -> None:
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    with _connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS presets (
                command TEXT NOT NULL,
                name TEXT NOT NULL,
                options_json TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                PRIMARY KEY (command, name)
            )
            """
        )


def list_presets(db_path: Path, command: str) -> list:
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT name, options_json FROM presets WHERE command = ? ORDER BY name",
            (command,),
        ).fetchall()
    return [{"name": name, "options": json.loads(options_json)} for name, options_json in rows]


def get_preset(db_path: Path, command: str, name: str):
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT options_json FROM presets WHERE command = ? AND name = ?",
            (command, name),
        ).fetchone()
    return json.loads(row[0]) if row else None


def save_preset(db_path: Path, command: str, name: str, options: dict, overwrite: bool = False) -> None:
    with _connect(db_path) as conn:
        exists = conn.execute(
            "SELECT 1 FROM presets WHERE command = ? AND name = ?", (command, name)
        ).fetchone()
        if exists and not overwrite:
            raise PresetExistsError(f"Preset '{name}' already exists for '{command}'")
        conn.execute(
            "INSERT OR REPLACE INTO presets (command, name, options_json) VALUES (?, ?, ?)",
            (command, name, json.dumps(options)),
        )


def delete_preset(db_path: Path, command: str, name: str) -> None:
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM presets WHERE command = ? AND name = ?", (command, name))
