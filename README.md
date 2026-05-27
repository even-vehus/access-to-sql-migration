# Migration Tools

Migration-only toolkit for extracting schema and data from Microsoft Access (`.accdb` / `.mdb`) and generating SQL-ready artifacts.

This repo intentionally excludes frontend/backend application code.

## Folder Structure

```text
migration-tools/
  access_databases/      # Put source Access files here
  migration/             # Extraction script
    extract_access_db.py
  migration_output/      # Generated SQL, CSV, and schema files (git-ignored)
  requirements.txt
```

## What It Generates

For each input database, outputs are created under `migration_output/<DatabaseName>/`:

- `01_create_tables.sql` - Table DDL (T-SQL)
- `02_foreign_keys.sql` - Foreign key constraints (if detected)
- `03_indexes.sql` - Index creation statements (if detected)
- `04_insert_data.sql` - Combined inserts in FK-friendly order
- `schema.json` - Extracted schema metadata
- `schema_summary.md` - Human-readable schema summary
- `csv_data/*.csv` - Raw exported table data
- `data_inserts/*.sql` - Per-table insert scripts

## Prerequisites (Windows)

1. Python 3.10+
2. Microsoft Access Database Engine / ODBC driver for Access
3. Optional: `pywin32` for richer ADOX schema extraction

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Usage

1. Place one or more Access files in `access_databases/` (any `.accdb`/`.mdb` names).
2. Run extraction:

```powershell
python .\migration\extract_access_db.py
```

Behavior:

- If there is one Access file, it is auto-selected.
- If there are multiple Access files, each one is processed into its own folder under `migration_output/<DatabaseName>/`.
- Use `--db-name` or `--db-path` to process just one database.

Or specify a file name explicitly:

```powershell
python .\migration\extract_access_db.py --db-name YourDatabase.accdb
```

Alternative explicit path:

```powershell
python .\migration\extract_access_db.py --db-path "C:\path\to\database.accdb"
```

Custom output directory:

```powershell
python .\migration\extract_access_db.py --db-name YourDatabase.accdb --output-dir .\migration_output\YourDatabase
```

If multiple databases are processed in one run, `--output-dir` is treated as the base output folder and each database gets its own subfolder under it.

## Notes

- By default ADOX is disabled to avoid COM/OLEDB stability issues.
- To enable ADOX on a machine where it is stable:

```powershell
$env:ACCESS_USE_ADOX = "1"
python .\migration\extract_access_db.py --db-name YourDatabase.accdb
```

- `migration_output/` is intentionally not version-controlled because it contains run-specific generated artifacts.
