# Web UI for bigmedia — Design Spec

## Purpose

`bigmedia` is currently a local CLI (`bigmedia sort ...`, `bigmedia dedupe
...`, etc, see [cli.py](../../../bigmedia/cli.py)). Non-technical team
members can't run it without a terminal and Python environment. This spec
adds a web front end so anyone on the team can run any `bigmedia`
operation from a browser: fill a form, upload `.xlsx` file(s), download
the result. Saved presets remove the need to re-type recurring option
combinations (e.g. `--name-column`, `--sheet`, `--categories`) each time.

## Non-goals

- No change to any business logic in `bigmedia/*.py` (classify, sort,
  dedupe, group, etc). The web layer only calls existing functions.
- No per-user accounts, no per-user preset ownership.
- No server-side folder-path input (CLI's folder-batch mode) — web only
  accepts direct file uploads.
- No permanent storage of uploaded/generated `.xlsx` content — client
  data must never persist beyond a single request/response cycle,
  matching the existing rule that real client data is never committed.

## Hosting & deployment

Deployed on **Railway** (PaaS), chosen for zero server-administration
overhead:

- Railway auto-detects a Python app from `requirements.txt` +
  `Procfile` and builds/runs it — no Dockerfile, no nginx/Caddy config,
  no systemd unit to write.
- Free HTTPS subdomain (`<project>.up.railway.app`) provided
  automatically.
- Deploy flow: code pushed to a GitHub repo (the `web/` addition to
  this repo), repo connected to a Railway project once via their
  dashboard, then every `git push` auto-redeploys.
- Cost: Railway's free trial credit is temporary; ongoing use requires
  their Hobby plan at **$5/month flat**. This is the cheapest tier that
  gives an always-on real URL without requiring server admin knowledge.
  (A same-office-network-only deployment on an existing PC would be
  free instead, but was ruled out since the team needs access from
  outside the office.)

### Persistent storage

SQLite file `presets.db` must survive redeploys, so it is placed on a
**Railway Volume** (a small persistent disk) mounted at `/data`.
Uploaded files and generated outputs are NOT persisted — they live only
in the container's ephemeral filesystem for the duration of a single
request and are deleted immediately after the response is sent (or on
error).

### Secrets & config

The shared team password (stored hashed, never plaintext) is set as a
Railway environment variable via their dashboard. No `.env` file is
committed to the repo.

## Auth

Single shared password, no individual accounts. On successful login, a
session cookie is set; all other routes require that cookie (redirect
to `/login` otherwise). This is sufficient because:

- Presets are team-shared, not personal — no need to know *who* is
  logged in.
- The tool touches no data more sensitive than clip lists already
  covered by the existing "never commit real client data" rule.

## Architecture

Single Python service:

```
web/
  main.py            FastAPI app: routes, session/auth middleware
  presets.py          SQLite-backed preset CRUD (table: command, name,
                        options_json, created_at)
  templates/            Jinja2 pages — one per bigmedia command, sharing
                          a form-layout partial
  static/                 minimal CSS/JS (vanilla or htmx, no build step)
requirements.txt
Procfile              `web: uvicorn web.main:app --host 0.0.0.0 --port $PORT`
```

The web layer imports and calls the existing `bigmedia/*.py` functions
directly (`sort_workbook()`, `dedupe_workbook()`, `group_duplicates_workbook()`,
`fu_grid_workbook()`, `compare_workbooks()`, `fix_getty_split()`,
`build_getty_id_report()`) — the same functions `cli.py` already wraps.
No subprocess invocation, no re-parsing of CLI args.

## Components

- **Command pages** — one form per subcommand, fields generated from
  that command's known options (mirrors the `argparse` definitions in
  `cli.py`): text inputs for things like `--name-column`, `--sheet`;
  numeric input for `--fps`; checkboxes for boolean flags like
  `--autofit`, `--no-group`; comma-separated text for list options like
  `--categories`.
- **File upload** — one or more `.xlsx` files via a standard browser
  file input.
- **Preset bar** — dropdown to load a saved preset into the current
  form (per command), a "Save as preset" button (prompts for a name;
  name collision asks overwrite-or-rename), and a way to delete a
  preset.
- **Presets storage** — SQLite table `presets(id, command, name,
  options_json, created_at)` on the Railway Volume.

## Data flow

1. User opens a command's page, optionally loads a preset (fills form
   fields from `options_json`), adjusts fields as needed.
2. User uploads one or more `.xlsx` files, submits the form.
3. Server validates every uploaded file has a `.xlsx` extension before
   calling any logic; rejects otherwise with an inline error naming the
   bad file(s).
4. For each valid file, the server calls the matching business function
   with the form's options, exactly as `cli.py`'s `process_one` does
   today for batch input.
5. Response: a single output file for a single upload, or a zip archive
   when multiple files were uploaded (mirrors the CLI's folder-batch
   naming, e.g. `EP21_master_sorted.xlsx`).
6. Uploaded and generated files are deleted from the ephemeral
   filesystem immediately after the response is sent (or immediately on
   error).

### Analyze step

Before running, the command form can call `POST /commands/{slug}/analyze`
with the uploaded files to inspect them read-only: it returns the header
union, per-sheet row counts, and (for commands with an analyzer, e.g.
Sort) suggested option values plus annotations like per-category clip
counts. The wizard uses this to populate the filename-column dropdown,
pre-tick category checkboxes, and surface warnings without blocking Run.
See `docs/superpowers/specs/2026-09-10-analyze-suggest-wizard-design.md`
for the full design, JSON shape, and precedence rules (analysis overrides
preset values).

## Error handling

- Each business-function call is wrapped; on exception, the form is
  re-rendered with an inline error banner containing the exception
  message (no stack trace shown to the user) and the offending
  filename, so a bad input never produces a silent failure or a raw
  500 page.
- Non-`.xlsx` uploads are rejected before any business logic runs.
- Preset name collisions prompt the user to overwrite or rename, never
  silently duplicate.
- Missing required fields (e.g. `--project-name` for `getty-ids`) are
  caught by standard HTML form validation plus a server-side check
  mirroring `argparse`'s `required=True`.

## Testing

- The existing `pytest` suite under `tests/` is untouched — no business
  logic moves or changes.
- New tests using FastAPI's `TestClient`, covering:
  - Auth gate: unauthenticated request to any command page redirects to
    `/login`; correct password sets a session cookie that then grants
    access.
  - Each command's happy path: valid form submission + a small fixture
    `.xlsx` (reusing patterns from `tests/conftest.py` where possible)
    produces a downloadable result with the expected content-type.
  - Preset CRUD: save, load (fields populate correctly), overwrite
    collision prompt, delete.
  - Bad upload rejection: non-`.xlsx` file produces the inline error,
    not a crash.

## Open questions / follow-ups (explicitly out of scope for this spec)

- Whether to eventually add per-user accounts if the team grows beyond
  trusting a shared password.
- Whether large uploads (many files at once) need a progress indicator
  — not addressed here; can be revisited once real usage patterns are
  known.
