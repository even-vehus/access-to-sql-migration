# Migration Tools

Migration toolkit for extracting schema and data from Microsoft Access (`.accdb` / `.mdb`) and generating SQL-ready artifacts.

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

### Option A: Run in Fabric Query Editor

Paste and execute the scripts in the order above.

### Option B: Run with sqlcmd

Verify `sqlcmd` is installed:

```powershell
sqlcmd -?
```

Run the generated scripts (replace `<YourDatabase>` with your extracted folder name):

```powershell
sqlcmd -S "<fabric-sql-endpoint>" -d "<database-name>" -U "<user>" -P "<password>" -i ".\migration_output\<YourDatabase>\01_create_tables.sql"
sqlcmd -S "<fabric-sql-endpoint>" -d "<database-name>" -U "<user>" -P "<password>" -i ".\migration_output\<YourDatabase>\04_insert_data.sql"
sqlcmd -S "<fabric-sql-endpoint>" -d "<database-name>" -U "<user>" -P "<password>" -i ".\migration_output\<YourDatabase>\03_indexes.sql"
```

If foreign keys were generated:

```powershell
sqlcmd -S "<fabric-sql-endpoint>" -d "<database-name>" -U "<user>" -P "<password>" -i ".\migration_output\<YourDatabase>\02_foreign_keys.sql"
```

## Notes

- By default ADOX is disabled to avoid COM/OLEDB stability issues.
- To enable ADOX on a machine where it is stable:

```powershell
$env:ACCESS_USE_ADOX = "1"
python .\migration\extract_access_db.py --db-name YourDatabase.accdb
```

- `migration_output/` is intentionally not version-controlled because it contains run-specific generated artifacts.
