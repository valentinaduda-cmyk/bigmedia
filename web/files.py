import zipfile
from pathlib import Path


def reject_non_xlsx(filenames):
    return [name for name in filenames if not name.lower().endswith(".xlsx")]


def parse_field(field, raw):
    if field.type == "checkbox":
        return bool(raw)
    if raw is None or raw == "" or raw == []:
        return None if field.type in ("list", "sheet_checklist", "number") else (field.default or "")
    if field.type in ("list", "sheet_checklist"):
        if isinstance(raw, list):
            parts = [p.strip() for p in raw if p.strip()]
        else:
            parts = [p.strip() for p in raw.split(",") if p.strip()]
        return parts or None
    if field.type == "number":
        num = float(raw)
        return int(num) if num.is_integer() else num
    return raw


def zip_files(paths, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in paths:
            zf.write(path, arcname=Path(path).name)
