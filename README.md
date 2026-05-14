# OPMM (Operation Plan Monitoring Matrix)

Django app for BSU Lipa operational plan monitoring: ingest Word or Excel monitor tables, store indicators and quarterly scores, and browse them in a performance viewer.

## Quick start

```powershell
cd path\to\SOPM
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

Open http://127.0.0.1:8000/ — you will be sent to the dashboard after login. Superusers are redirected to **Performance viewer** (`/performance-viewer/`).

## Sample data (local testing)

Generate three demo offices, PAPs, indicators, and performance rows (aligned with the performance viewer development areas):

```powershell
python manage.py seed_sample_data --reset --with-users
```

- `--reset` deletes **all** strategic rows (performance, indicators, strategic levels), removes prior **Demo —** offices and `demo.*` users, then reloads fixtures.
- Omit `--reset` only on a database that has **no** strategic data yet.
- `--with-users` creates `demo.central`, `demo.east`, and `demo.west` (password `demo12345`) with `first_name` matching each demo office for office-scoped dashboards.
- **Portable copy for another machine:** `core/fixtures/sample_strategic_data.json` (Demo-only snapshot). Regenerate with `python manage.py export_demo_fixture` after `seed_sample_data --reset`. Load elsewhere with `python manage.py loaddata sample_strategic_data` (see `core/fixtures/LOAD_INSTRUCTIONS.txt` for PK conflicts).

**Larger load test (~100 extra indicators, 400+ extra performance rows):**

```powershell
python manage.py seed_sample_data --reset --stress
# equivalent to:
python manage.py seed_sample_data --reset --bulk-indicators 100 --bulk-quarters 4
```

- `--bulk-indicators N` adds `N` synthetic KPIs split across the three demo offices (after the small curated set).
- `--bulk-quarters Q` (1–4, default 2) creates that many `PerformanceRecord` rows per bulk indicator (year 2026).
- `--stress` is a shortcut for `--bulk-indicators 100 --bulk-quarters 4` (~429 total performance rows with the curated seed).

**Even demo across the six development areas (100 indicators per area, ~80% met):**

```powershell
python manage.py seed_sample_data --reset --balanced-only --balanced-per-area 100
```

- Loads **only** the three demo offices plus **600** indicators (six OPMM areas × 100), each with one `PerformanceRecord` in FY **2026 Q1** (change quarter with `--balanced-quarter` 2–4).
- About **80% met** in five areas; **Sustainability** uses one fewer met row so the sidebar is not perfectly flat.
- `--balanced-met-pct` (default 80) rounds the met count per area; combine with `--with-users` as usual.
- Without `--balanced-only`, `--balanced-per-area 100` **adds** 600 rows on top of the curated sample (area totals will be higher than 100).

Use a **superuser** for the performance viewer, charts, and upload-as-admin behavior.

## Main URLs

| Path | Purpose |
|------|---------|
| `/` | Dashboard (alias of `/dashboard/`) |
| `/accounts/login/` | Login |
| `/upload/` | Upload a `.docx` or `.xlsx` monitor file |
| `/performance-viewer/` | Admin summary (superuser) |
| `/users/` | User / office registration (superuser) |
| `/admin/` | Django admin |

## Office ↔ user linking

Non–superusers only see data for their office. The app resolves the office by matching **`User.first_name`** to **`Office.name`** (see `_office_for_user` in `core/views.py`). When an admin registers a user, the full office name is stored in `first_name` and a matching `Office` row is created — keep those names identical.

## Blueprint (.docx / .xlsx) format

**Word:** first **table**, data rows from **row index 3** (fourth row).

**Excel:** first **worksheet**, same columns in **A–E**, data from **row 4** (rows 1–3 are headers / period hints).

Columns for both: outcome, strategy, PAP, indicator text, accomplishment text. Quarter and year can be chosen on the upload form or inferred from title / top rows (see `core/services.py`).

## Configuration (environment variables)

| Variable | Purpose |
|----------|---------|
| `DJANGO_SECRET_KEY` | Secret key; **required** in production |
| `DJANGO_DEBUG` | `true` / `false` (default `true` if unset) |
| `DJANGO_ALLOWED_HOSTS` | Comma-separated hostnames (default `localhost,127.0.0.1`) |
| `SOPM_ENABLE_FULL_DATABASE_CLEAR` | When `DEBUG` is false, set to `true` only if superusers may use **Reset Data** (POST to `/clear-data/`) |

Office-only **Clear office data** (`/clear-office-data/`) remains available to signed-in staff via POST; it does not use `SOPM_ENABLE_FULL_DATABASE_CLEAR`.

## Tests

```powershell
python manage.py test core.tests
```
