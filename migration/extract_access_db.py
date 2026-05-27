"""
Access Database Extractor
Extracts schema, data, relationships, and queries from an Access .accdb file
and generates Microsoft Fabric SQL (T-SQL) migration scripts.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
from collections import deque
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pyodbc

try:
    import win32com.client

    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    default_access_dir = repo_root / "access_databases"

    parser = argparse.ArgumentParser(
        description="Extract Access (.accdb) and generate SQL/CSV files for Fabric SQL import."
    )
    parser.add_argument(
        "--access-dir",
        default=str(default_access_dir),
        help="Directory containing Access databases (default: ./access_databases)",
    )
    parser.add_argument(
        "--db-name",
        default="Northwind.accdb",
        help="Access file name inside --access-dir when --db-path is not provided",
    )
    parser.add_argument(
        "--db-path",
        default=None,
        help="Full path to a .accdb file (overrides --access-dir and --db-name)",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Output directory for generated files (default: ./migration_output/<db_name_without_ext>)",
    )
    return parser.parse_args()


def get_connection(db_path: Path):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    return pyodbc.connect(conn_str)


def get_adox_schema(db_path: Path) -> dict:
    """Use ADOX via win32com to get richer schema info (AutoNumber, PKs, FKs, field types)."""
    # ADOX can hang on some Office/ACE installations. Skip it by default and
    # allow opt-in when a machine has stable COM/OLEDB behavior.
    if os.getenv("ACCESS_USE_ADOX", "0") != "1":
        print("  ADOX skipped (set ACCESS_USE_ADOX=1 to enable)")
        return {}
    if not HAS_WIN32COM:
        return {}
    schema = {}
    try:
        cat = win32com.client.Dispatch("ADOX.Catalog")
        cat.ActiveConnection = f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db_path};"
        for table in cat.Tables:
            if table.Name.startswith("MSys"):
                continue
            tbl_info = {
                "columns": {},
                "primary_keys": [],
                "foreign_keys": [],
                "indexes": [],
            }
            for col in table.Columns:
                props = {}
                try:
                    for prop in col.Properties:
                        try:
                            props[prop.Name] = prop.Value
                        except Exception:
                            pass
                except Exception:
                    pass
                tbl_info["columns"][col.Name] = {
                    "adox_type": col.Type,
                    "adox_size": col.DefinedSize,
                    "is_autonumber": props.get("Autoincrement", False),
                    "is_required": not props.get("Nullable", True),
                    "default_value": props.get("Default", None),
                    "description": props.get("Description", None),
                }
            try:
                for key in table.Keys:
                    if key.Type == 1:  # Primary
                        for col in key.Columns:
                            tbl_info["primary_keys"].append(col.Name)
                    elif key.Type == 2:  # Foreign
                        fk_cols = [col.Name for col in key.Columns]
                        rel_table = key.RelatedTable
                        rel_cols = [col.RelatedColumn for col in key.Columns]
                        for fk_col, rel_col in zip(fk_cols, rel_cols):
                            tbl_info["foreign_keys"].append(
                                {
                                    "fk_name": key.Name,
                                    "fk_column": fk_col,
                                    "pk_table": rel_table,
                                    "pk_column": rel_col,
                                }
                            )
            except Exception:
                pass

            try:
                for idx in table.Indexes:
                    idx_cols = [col.Name for col in idx.Columns]
                    tbl_info["indexes"].append(
                        {
                            "name": idx.Name,
                            "unique": idx.Unique,
                            "primary": idx.PrimaryKey,
                            "columns": idx_cols,
                        }
                    )
            except Exception:
                pass
            schema[table.Name] = tbl_info
    except Exception as e:
        print(f"  [ADOX warning] {e}")
    return schema


def sanitize_identifier(name: str) -> str:
    return f"[{name}]"


def escape_sql_string(value) -> str:
    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, datetime):
        return f"'{value.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}'"
    if isinstance(value, date):
        return f"'{value.strftime('%Y-%m-%d')}'"
    if isinstance(value, bytes):
        return f"0x{value.hex()}"
    escaped = str(value).replace("'", "''")
    return f"N'{escaped}'"


def get_tables(conn) -> list[str]:
    cursor = conn.cursor()
    return [
        row.table_name
        for row in cursor.tables(tableType="TABLE")
        if not row.table_name.startswith("MSys")
    ]


def get_column_info(conn, table_name: str) -> list[dict]:
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM [{table_name}] WHERE 1=0")
    columns = []
    for i, desc in enumerate(cursor.description):
        name, type_code, _display_size, internal_size, precision, scale, null_ok = desc
        if type_code == str:
            type_name = "NVARCHAR"
        elif type_code == int:
            type_name = "LONG"
        elif type_code == float:
            type_name = "DOUBLE"
        elif type_code == bool:
            type_name = "BIT"
        elif type_code == datetime:
            type_name = "DATETIME"
        elif type_code == date:
            type_name = "DATE"
        elif type_code == bytes:
            type_name = "LONGBINARY"
        elif type_code == Decimal:
            type_name = "DECIMAL"
        else:
            type_name = "NVARCHAR"
        columns.append(
            {
                "name": name,
                "type_name": type_name,
                "column_size": internal_size,
                "decimal_digits": scale if scale is not None else 0,
                "nullable": bool(null_ok),
                "column_def": None,
                "ordinal_position": i + 1,
                "_type_code": type_code,
                "_precision": precision,
            }
        )
    return columns


def get_primary_keys(conn, table_name: str) -> list[str]:
    cursor = conn.cursor()
    pks = []
    try:
        for row in cursor.primaryKeys(table=table_name):
            pks.append((row.key_seq, row.column_name))
        pks.sort(key=lambda x: x[0])
        return [col for _, col in pks]
    except Exception:
        return []


def get_foreign_keys(conn, table_name: str) -> list[dict]:
    cursor = conn.cursor()
    fks = []
    try:
        for row in cursor.foreignKeys(foreignTable=table_name):
            fks.append(
                {
                    "fk_column": row.fkcolumn_name,
                    "pk_table": row.pktable_name,
                    "pk_column": row.pkcolumn_name,
                    "fk_name": row.fk_name,
                }
            )
    except Exception:
        pass
    return fks


def get_indexes(conn, table_name: str) -> list[dict]:
    cursor = conn.cursor()
    indexes = {}
    try:
        for row in cursor.statistics(table=table_name):
            if row.index_name and not row.index_name.startswith("MSys"):
                idx_name = row.index_name
                if idx_name not in indexes:
                    indexes[idx_name] = {
                        "name": idx_name,
                        "unique": row.non_unique == 0,
                        "columns": [],
                    }
                indexes[idx_name]["columns"].append((row.ordinal_position, row.column_name))
        for idx in indexes.values():
            idx["columns"] = [
                col for _, col in sorted(idx["columns"], key=lambda x: x[0] or 0)
            ]
    except Exception:
        pass
    return list(indexes.values())


def get_table_data(conn, table_name: str):
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM [{table_name}]")
    col_descs = cursor.description
    rows = cursor.fetchall()
    return col_descs, rows


def access_type_to_tsql(col: dict) -> str:
    type_code = col.get("_type_code")
    size = col.get("column_size")
    precision = col.get("_precision") or col.get("column_size")
    scale = col.get("decimal_digits")

    if type_code == str:
        if size and 0 < size <= 4000:
            return f"NVARCHAR({size})"
        return "NVARCHAR(MAX)"
    if type_code == int:
        return "INT"
    if type_code == float:
        return "FLOAT"
    if type_code == bool:
        return "BIT"
    if type_code == datetime:
        return "DATETIME2"
    if type_code == date:
        return "DATE"
    if type_code == bytes:
        return "VARBINARY(MAX)"
    if type_code == Decimal:
        p = precision if precision else 18
        s = scale if scale is not None else 4
        return f"DECIMAL({p},{s})"

    type_name = col.get("type_name", "NVARCHAR").upper()
    base_type = re.sub(r"\(.*\)", "", type_name).strip()

    if base_type in ("VARCHAR", "NVARCHAR", "CHAR", "NCHAR", "WVARCHAR", "TEXT"):
        if size and 0 < size <= 4000:
            return f"NVARCHAR({size})"
        return "NVARCHAR(MAX)"
    if base_type in ("LONGCHAR", "MEMO", "LONGTEXT"):
        return "NVARCHAR(MAX)"
    if base_type in ("COUNTER", "AUTOINCREMENT"):
        return "INT IDENTITY(1,1)"
    if base_type in ("LONG", "INTEGER"):
        return "INT"
    if base_type in ("SHORT", "SMALLINT"):
        return "SMALLINT"
    if base_type == "BYTE":
        return "TINYINT"
    if base_type in ("DOUBLE", "FLOAT", "IEEEDOUBLE"):
        return "FLOAT"
    if base_type in ("SINGLE", "IEEESINGLE", "REAL"):
        return "REAL"
    if base_type in ("CURRENCY", "MONEY"):
        return "DECIMAL(19,4)"
    if base_type in ("DECIMAL", "NUMERIC"):
        p = precision if precision else 18
        s = scale if scale is not None else 4
        return f"DECIMAL({p},{s})"
    if base_type in ("DATETIME", "TIMESTAMP"):
        return "DATETIME2"
    if base_type == "DATE":
        return "DATE"
    if base_type == "TIME":
        return "TIME"
    if base_type in ("BIT", "YESNO", "BOOLEAN"):
        return "BIT"
    if base_type in ("BINARY", "VARBINARY", "LONGBINARY", "IMAGE"):
        return "VARBINARY(MAX)"
    if base_type == "GUID":
        return "UNIQUEIDENTIFIER"
    if base_type == "BIGINT":
        return "BIGINT"
    return f"NVARCHAR(255) /* original: {type_name} */"


def build_create_table_sql(
    table_name: str, columns: list[dict], primary_keys: list[str], adox_table: dict | None = None
) -> str:
    col_lines = []

    for col in columns:
        adox_col = (adox_table or {}).get(col["name"], {})
        if adox_col.get("is_autonumber"):
            tsql_type = "INT IDENTITY(1,1)"
            nullable = ""
        else:
            tsql_type = access_type_to_tsql(col)
            nullable = "" if col["nullable"] else " NOT NULL"
            if "IDENTITY" in tsql_type:
                nullable = ""
        col_lines.append(f"    {sanitize_identifier(col['name'])} {tsql_type}{nullable}")

    if primary_keys:
        pk_cols = ", ".join(sanitize_identifier(pk) for pk in primary_keys)
        col_lines.append(f"    CONSTRAINT [PK_{table_name}] PRIMARY KEY ({pk_cols})")

    return "\n".join(
        [
            f"CREATE TABLE {sanitize_identifier(table_name)} (",
            ",\n".join(col_lines),
            ");",
        ]
    )


def build_insert_sql(table_name: str, col_descs, rows, has_identity: bool = False) -> list[str]:
    if not rows:
        return []

    col_names = ", ".join(sanitize_identifier(d[0]) for d in col_descs)
    statements = []

    inserts = []
    for row in rows:
        values = ", ".join(escape_sql_string(v) for v in row)
        inserts.append(f"    ({values})")

    prefix = []
    suffix = []
    if has_identity:
        prefix = [f"SET IDENTITY_INSERT {sanitize_identifier(table_name)} ON;"]
        suffix = [f"SET IDENTITY_INSERT {sanitize_identifier(table_name)} OFF;"]

    batch_size = 1000
    for i in range(0, len(inserts), batch_size):
        batch = inserts[i : i + batch_size]
        sql = (
            f"INSERT INTO {sanitize_identifier(table_name)} ({col_names})\nVALUES\n"
            + ",\n".join(batch)
            + ";"
        )
        statements.append(sql)

    return prefix + statements + suffix


def export_table_csv(table_name: str, col_descs, rows, output_dir: Path):
    csv_path = output_dir / "csv_data" / f"{table_name}.csv"
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([d[0] for d in col_descs])
        for row in rows:
            writer.writerow(
                [
                    v.isoformat() if isinstance(v, (datetime, date)) else str(v) if v is not None else ""
                    for v in row
                ]
            )


def main():
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    access_dir = Path(args.access_dir).resolve()
    access_dir.mkdir(parents=True, exist_ok=True)

    if args.db_path:
        db_path = Path(args.db_path).resolve()
    else:
        db_path = (access_dir / args.db_name).resolve()

    if args.output_dir:
        output_dir = Path(args.output_dir).resolve()
    else:
        output_dir = (repo_root / "migration_output" / db_path.stem).resolve()

    if not db_path.exists():
        raise FileNotFoundError(
            f"Access file not found: {db_path}\n"
            f"Put your .accdb in: {access_dir}\n"
            "or pass --db-path explicitly."
        )

    print(f"Connecting to: {db_path}")
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "csv_data").mkdir(exist_ok=True)

    print("Reading ADOX schema...")
    adox_schema = get_adox_schema(db_path)
    if adox_schema:
        print(f"  ADOX schema loaded for {len(adox_schema)} tables")
    else:
        print("  ADOX unavailable, falling back to pyodbc types only")

    conn = get_connection(db_path)

    tables = get_tables(conn)
    print(f"Found {len(tables)} tables: {tables}")

    schema_info = {}
    ddl_statements = []
    fk_statements = []
    index_statements = []
    all_inserts = {}

    ddl_statements.append(
        f"-- Microsoft Fabric SQL Migration Script\n"
        f"-- Generated: {datetime.now().isoformat()}\n"
        f"-- Source: {db_path.name}\n"
        f"-- Target: Microsoft Fabric SQL Database (T-SQL)\n"
    )

    for table_name in tables:
        print(f"  Processing table: {table_name}")
        columns = get_column_info(conn, table_name)
        adox_table = adox_schema.get(table_name, {})
        primary_keys = adox_table.get("primary_keys") or get_primary_keys(conn, table_name)
        foreign_keys = adox_table.get("foreign_keys") or get_foreign_keys(conn, table_name)
        indexes = adox_table.get("indexes") or get_indexes(conn, table_name)
        col_descs, rows = get_table_data(conn, table_name)

        schema_info[table_name] = {
            "columns": columns,
            "primary_keys": primary_keys,
            "foreign_keys": foreign_keys,
            "indexes": indexes,
            "row_count": len(rows),
        }

        create_sql = build_create_table_sql(table_name, columns, primary_keys, adox_table.get("columns"))
        ddl_statements.append(f"-- Table: {table_name} ({len(rows)} rows)")
        ddl_statements.append(create_sql)
        ddl_statements.append("")

        for fk in foreign_keys:
            constraint_name = fk["fk_name"] or f"FK_{table_name}_{fk['fk_column']}"
            fk_sql = (
                f"ALTER TABLE {sanitize_identifier(table_name)}\n"
                f"  ADD CONSTRAINT {sanitize_identifier(constraint_name)}\n"
                f"  FOREIGN KEY ({sanitize_identifier(fk['fk_column'])})\n"
                f"  REFERENCES {sanitize_identifier(fk['pk_table'])} ({sanitize_identifier(fk['pk_column'])});"
            )
            fk_statements.append(fk_sql)

        pk_index_names = {f"PK_{table_name}"}
        for idx in indexes:
            if idx["name"] in pk_index_names:
                continue
            if not idx["columns"]:
                continue
            unique_kw = "UNIQUE " if idx["unique"] else ""
            cols = ", ".join(sanitize_identifier(c) for c in idx["columns"])
            idx_sql = (
                f"CREATE {unique_kw}INDEX {sanitize_identifier(idx['name'])}\n"
                f"  ON {sanitize_identifier(table_name)} ({cols});"
            )
            index_statements.append(idx_sql)

        adox_cols = adox_table.get("columns", {})
        has_identity = any(c.get("is_autonumber") for c in adox_cols.values())
        insert_sqls = build_insert_sql(table_name, col_descs, rows, has_identity)
        if insert_sqls:
            all_inserts[table_name] = insert_sqls

        export_table_csv(table_name, col_descs, rows, output_dir)

    conn.close()

    ddl_path = output_dir / "01_create_tables.sql"
    with open(ddl_path, "w", encoding="utf-8") as f:
        f.write("\n".join(ddl_statements))
    print(f"  Written: {ddl_path}")

    if fk_statements:
        fk_path = output_dir / "02_foreign_keys.sql"
        with open(fk_path, "w", encoding="utf-8") as f:
            f.write("-- Foreign Key Constraints\n\n")
            f.write("\n\n".join(fk_statements))
        print(f"  Written: {fk_path}")

    if index_statements:
        idx_path = output_dir / "03_indexes.sql"
        with open(idx_path, "w", encoding="utf-8") as f:
            f.write("-- Indexes\n\n")
            f.write("\n\n".join(index_statements))
        print(f"  Written: {idx_path}")

    data_dir = output_dir / "data_inserts"
    data_dir.mkdir(exist_ok=True)
    for table_name, inserts in all_inserts.items():
        data_path = data_dir / f"{table_name}.sql"
        with open(data_path, "w", encoding="utf-8") as f:
            f.write(f"-- Data for table: {table_name}\n")
            f.write(f"-- Row count: {schema_info[table_name]['row_count']}\n\n")
            f.write("\n\n".join(inserts))
        print(f"  Written: {data_path}")

    deps: dict[str, set[str]] = {t: set() for t in tables}
    for table_name in tables:
        for fk in schema_info[table_name]["foreign_keys"]:
            parent = fk["pk_table"]
            if parent in deps and parent != table_name:
                deps[table_name].add(parent)

    in_degree = {t: len(d) for t, d in deps.items()}
    queue = deque(t for t, d in in_degree.items() if d == 0)
    ordered_tables = []
    while queue:
        node = queue.popleft()
        ordered_tables.append(node)
        for t in tables:
            if node in deps[t]:
                deps[t].discard(node)
                in_degree[t] -= 1
                if in_degree[t] == 0:
                    queue.append(t)
    for t in tables:
        if t not in ordered_tables:
            ordered_tables.append(t)

    combined_data_path = output_dir / "04_insert_data.sql"
    with open(combined_data_path, "w", encoding="utf-8") as f:
        f.write("-- Combined Data Inserts (FK dependency order)\n")
        f.write("-- Insert order: " + " -> ".join(ordered_tables) + "\n\n")
        for table_name in ordered_tables:
            if table_name in all_inserts:
                f.write(f"-- === {table_name} ===\n")
                f.write("\n".join(all_inserts[table_name]))
                f.write("\n\n")
    print(f"  Written: {combined_data_path}")

    json_path = output_dir / "schema.json"

    def json_serializer(obj):
        if isinstance(obj, type):
            return obj.__name__
        raise TypeError(f"Object of type {type(obj)} is not JSON serializable")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(schema_info, f, indent=2, default=json_serializer)
    print(f"  Written: {json_path}")

    erd_path = output_dir / "schema_summary.md"
    with open(erd_path, "w", encoding="utf-8") as f:
        f.write("# Database Schema Summary\n\n")
        f.write(f"**Source:** {db_path.name}  \n")
        f.write(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M')}  \n\n")
        f.write(f"## Tables ({len(tables)})\n\n")
        for table_name in tables:
            info = schema_info[table_name]
            f.write(f"### {table_name} ({info['row_count']} rows)\n\n")
            f.write("| Column | Type (Access) | T-SQL Type | PK | Nullable |\n")
            f.write("|--------|--------------|------------|----|---------|\n")
            for col in info["columns"]:
                is_pk = "x" if col["name"] in info["primary_keys"] else ""
                nullable = "x" if col["nullable"] else ""
                tsql_type = access_type_to_tsql(col)
                f.write(
                    f"| {col['name']} | {col['type_name']} | {tsql_type} | {is_pk} | {nullable} |\n"
                )
            if info["foreign_keys"]:
                f.write("\n**Foreign Keys:**\n\n")
                for fk in info["foreign_keys"]:
                    f.write(f"- {fk['fk_column']} -> {fk['pk_table']}.{fk['pk_column']}\n")
            f.write("\n")

    print(f"  Written: {erd_path}")

    print("\n=== EXTRACTION COMPLETE ===")
    print(f"Output directory: {output_dir}")
    print(f"Tables extracted: {len(tables)}")
    total_rows = sum(info["row_count"] for info in schema_info.values())
    print(f"Total rows: {total_rows}")
    print("\nFiles generated:")
    for p in sorted(output_dir.rglob("*")):
        if p.is_file():
            size = p.stat().st_size
            print(f"  {p.relative_to(output_dir)}  ({size:,} bytes)")


if __name__ == "__main__":
    main()
