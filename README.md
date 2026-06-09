# Migration Tools

Migration toolkit for extracting schema, data, forms, and VBA from Microsoft Access (`.accdb` / `.mdb`), generating SQL-ready artifacts for Microsoft Fabric SQL Database, and producing an app-spec bundle for app modernization.

## Folder Structure

```text
access-to-sql-migration/
  access_databases/             # Put Access files here (git-ignored except .gitkeep)
  migration/
    extract_access_db.py        # Step 1: schema + data + queries → SQL/CSV/JSON
  extract/
    extract_forms_vba.py        # Step 2: forms, reports, VBA, macros (via Access COM)
  generators/
    inspect_artifacts.py        # Step 3: schema.json + forms_vba.json → app_spec.*
    configs/                    # Optional per-database override configs
  deploy/
    push_to_fabric.py           # Deploy generated SQL to Fabric via sqlcmd
  tests/                        # pytest suite for the scripts above
  migration_output/             # Generated artifacts, per database (git-ignored)
  requirements.txt              # Runtime dependencies
  requirements-dev.txt          # Test dependencies
  MIGRATION_WORKFLOW.md         # End-to-end modernization workflow (factory → forge)
```

## Pipeline

The tools run as a three-step pipeline. All artifacts for a database land in `migration_output/<DatabaseName>/`.

| Step | Script | Produces |
|------|--------|----------|
| 1. Extract schema & data | `migration/extract_access_db.py` | T-SQL DDL, inserts, CSVs, `schema.json` |
| 2. Extract forms & VBA | `extract/extract_forms_vba.py` | `forms_vba.json`, `forms_vba_summary.md` |
| 3. Generate app spec | `generators/inspect_artifacts.py` | `app_spec.json`, `app_spec.md` |

Steps 1 and 2 are independent extractors; step 3 consumes the JSON from both. For a plain SQL-to-Fabric migration you only need step 1. Steps 2–3 feed the app-modernization workflow in [MIGRATION_WORKFLOW.md](MIGRATION_WORKFLOW.md).

## What It Generates

Outputs are created under `migration_output/<DatabaseName>/`.

**Step 1 — `extract_access_db.py`:**

- `01_create_tables.sql` - Table DDL (T-SQL)
- `02_foreign_keys.sql` - Foreign key constraints (if detected)
- `03_indexes.sql` - Index creation statements (if detected)
- `04_insert_data.sql` - Combined inserts in FK-friendly order
- `05_views.sql` - Saved Access SELECT/UNION queries converted to T-SQL views (queries referencing form controls are commented out for manual handling)
- `06_action_queries.sql` - Action queries (UPDATE/DELETE/APPEND/MAKE-TABLE) preserved as reference comments; migrate manually
- `schema.json` - Extracted schema metadata
- `schema_summary.md` - Human-readable schema summary
- `csv_data/*.csv` - Raw exported table data
- `data_inserts/*.sql` - Per-table insert scripts

**Step 2 — `extract_forms_vba.py`:**

- `forms_vba.json` - Forms, reports, VBA modules, and macros, plus complexity metrics
- `forms_vba_summary.md` - Human-readable summary of the above

**Step 3 — `inspect_artifacts.py`:**

- `app_spec.json` - Compact machine-readable specification for code generation
- `app_spec.md` - Human-readable overview (entities → C# classes, forms → routes, VBA classification)

## Prerequisites (Windows)

1. Python 3.10+
2. Microsoft Access Database Engine / ODBC driver for Access (used by step 1)
3. Microsoft Access (the desktop application) is required for step 2, which drives Access via COM automation
4. To extract VBA source in step 2, enable: File → Options → Trust Center → Trust Center Settings → Macro Settings → "Trust access to the VBA project object model"

`pywin32` (in `requirements.txt`) is required for step 2 and for richer ADOX schema extraction in step 1.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Usage

Place one or more Access files in `access_databases/` (any `.accdb`/`.mdb` names).

Selection behavior is the same across all three scripts:

- If there is one Access file / output folder, it is auto-selected.
- If there are multiple, each is processed into its own folder under `migration_output/<DatabaseName>/`.
- Use `--db-name` (or `--db-path` on the extractors) to process just one.

### Step 1 — Extract schema & data

```powershell
python .\migration\extract_access_db.py
```

Specify a file name, an explicit path, or a custom output directory:

```powershell
python .\migration\extract_access_db.py --db-name YourDatabase.accdb
python .\migration\extract_access_db.py --db-path "C:\path\to\database.accdb"
python .\migration\extract_access_db.py --db-name YourDatabase.accdb --output-dir .\migration_output\YourDatabase
```

If multiple databases are processed in one run, `--output-dir` is treated as the base output folder and each database gets its own subfolder under it.

### Step 2 — Extract forms & VBA (optional)

Requires the Access desktop app. Drives it via COM automation to pull out UI and code that ODBC/ADOX cannot see.

```powershell
python .\extract\extract_forms_vba.py --db-name YourDatabase.accdb
```

Use `--password` for a password-protected database. Writes into the same `migration_output/<DatabaseName>/` folder as step 1.

### Step 3 — Generate app spec (optional)

Reads `schema.json` + `forms_vba.json` from a `migration_output/<DatabaseName>/` folder and writes `app_spec.json` + `app_spec.md`.

```powershell
python .\generators\inspect_artifacts.py --db-name YourDatabase
```

Genuine product decisions no heuristic can infer (which entities are "v1 core", the app title, C# class names, custom routes) can be pinned in an optional override config — pass `--config path/to/overrides.json`, or drop `generators/configs/<db-name>.json` to have it auto-loaded.

## Recreate Tables in Fabric SQL Database

This project is designed for **Microsoft Fabric SQL Database**.

1. In Fabric, create a SQL Database in your workspace.
2. Open its SQL connection endpoint details.
3. Pick the matching extracted folder under `migration_output/<DatabaseName>/`.
4. Run scripts in this order:

- `01_create_tables.sql` (create schema)
- `04_insert_data.sql` (load rows)
- `03_indexes.sql` (create indexes)
- `02_foreign_keys.sql` (optional, run last if generated)

### Deploy to Fabric

`push_to_fabric.py` can run the generated scripts to push data to a selected Fabric SQL database.

Run from the repo root using the project venv Python (make sure you are signed in with an Entra ID that has access to the SQL Database you are pushing to):

```powershell
python deploy/push_to_fabric.py `
  --connection-string "<connection-string>" `
  --db-name <local-database-name> `
```

Options of note:
- `--include-views`: also run `05_views.sql` (only if you've reviewed/translated Access SQL to T-SQL)
- `--skip-fk`: skip `02_foreign_keys.sql`
- `--auth password` (fallback): uses `--username` and `FABRIC_SQL_PASSWORD` env var; avoid when possible

See `deploy/push_to_fabric.py` for full flags and behavior.

## Testing

```powershell
pip install -r requirements-dev.txt
python -m pytest
```

## Further Reading

[MIGRATION_WORKFLOW.md](MIGRATION_WORKFLOW.md) describes the end-to-end app-modernization workflow these scripts feed into — turning the `app_spec` bundle into a modern web app.
