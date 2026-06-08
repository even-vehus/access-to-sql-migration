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
import time
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
        default=None,
        help=(
            "Access file name inside --access-dir when --db-path is not provided. "
            "If omitted, the script auto-selects the only .accdb/.mdb in --access-dir, "
            "or processes all of them into separate output folders."
        ),
    )
    parser.add_argument(
        "--db-path",
        default=None,
        help="Full path to a .accdb file (overrides --access-dir and --db-name)",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help=(
            "Output directory for generated files (default: ./migration_output/<db_name_without_ext>). "
            "When multiple databases are processed, each database is written to a subfolder under this path."
        ),
    )
    return parser.parse_args()


def get_connection(db_path: Path):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={db_path};"
    )
    last_exc: Exception | None = None
    for attempt in range(1, 4):
        try:
            return pyodbc.connect(conn_str)
        except (pyodbc.Error, SystemError) as exc:
            last_exc = exc
            if attempt < 3:
                wait_seconds = 0.5 * attempt
                print(
                    f"  [connect retry {attempt}/3] Access driver was busy/unavailable; retrying in {wait_seconds:.1f}s..."
                )
                time.sleep(wait_seconds)

    raise RuntimeError(
        f"Failed to connect to Access database after 3 attempts: {db_path}\n"
        "Close Access/Office processes locking the file and try again."
    ) from last_exc


def get_adox_schema(db_path: Path) -> dict:
    """Use ADOX via win32com to get richer schema info (AutoNumber, PKs, FKs, field types)."""
    # ADOX can hang on some Office/ACE installations. Enabled by default;
    # set ACCESS_USE_ADOX=0 to disable if you experience hangs.
    if os.getenv("ACCESS_USE_ADOX", "1") != "1":
        print("  ADOX skipped (set ACCESS_USE_ADOX=0 to disable)")
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


DAO_QUERY_TYPE_NAMES: dict[int, str] = {
    0: "SELECT",
    16: "UNION",
    32: "CROSSTAB",
    48: "DELETE",
    64: "UPDATE",
    80: "APPEND",
    96: "MAKE-TABLE",
    112: "DDL",
    128: "PASS-THROUGH",
    160: "PROCEDURE",
    240: "DATA-DEFINITION",
}

VIEW_QUERY_TYPES = {0, 16}  # Types that can become SQL views

# ── Access → T-SQL conversion helpers ────────────────────────────────────────

_ACCESS_DATEADD_MAP: dict[str, str] = {
    "yyyy": "year",
    "q": "quarter",
    "m": "month",
    "y": "dayofyear",
    "d": "day",
    "w": "weekday",
    "ww": "week",
    "h": "hour",
    "n": "minute",
    "s": "second",
}


def _strip_order_by_if_no_top(sql: str) -> str:
    """Remove trailing ORDER BY when the SELECT has no TOP clause (invalid in T-SQL views)."""
    sel = re.search(r"\bSELECT\b", sql, re.IGNORECASE)
    if sel:
        after = sql[sel.end() : sel.end() + 40]
        if re.match(r"\s+TOP\b", after, re.IGNORECASE):
            return sql  # TOP present — ORDER BY is valid
    depth, last_pos, i, n = 0, -1, 0, len(sql)
    while i < n:
        c = sql[i]
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif depth == 0:
            if re.match(r"ORDER\s+BY\b", sql[i:], re.IGNORECASE):
                last_pos = i
        i += 1
    return sql[:last_pos].rstrip() if last_pos >= 0 else sql


def _convert_access_operators(sql: str) -> str:
    """
    Character-level scan to convert Access-specific operators:
      - Double-quoted literals "..." → single-quoted T-SQL strings 'safe''quote'
      - Access concatenation operator & → + (only outside string literals)
    """
    out: list[str] = []
    i, n = 0, len(sql)
    while i < n:
        c = sql[i]
        if c == '"':
            # Read Access double-quoted string literal → convert to single-quoted
            j = i + 1
            buf: list[str] = []
            while j < n and sql[j] != '"':
                ch = sql[j]
                if ch == "'":
                    buf.append("''")  # escape single quotes for T-SQL
                else:
                    buf.append(ch)
                j += 1
            out.append("'" + "".join(buf) + "'")
            i = j + 1
        elif c == "'":
            # Pass through existing single-quoted strings unchanged (handle '' escapes)
            out.append(c)
            j = i + 1
            while j < n:
                out.append(sql[j])
                if sql[j] == "'":
                    j += 1
                    if j < n and sql[j] == "'":
                        out.append("'")
                        j += 1
                    else:
                        break
                else:
                    j += 1
            i = j
        elif c == '&':
            out.append(' + ')
            i += 1
        else:
            out.append(c)
            i += 1
    return "".join(out)


def access_sql_to_tsql(sql: str) -> tuple[str, list[str]]:
    """
    Best-effort conversion of Access SQL to T-SQL for SQL Server / Fabric SQL.
    Returns (converted_sql, list_of_notes).
    """
    notes: list[str] = []

    # 1. Replace Access ! field-separator with .  (preserve != operator)
    sql = re.sub(r"!(?!=)", ".", sql)

    # 2. Convert double-quoted string literals → single-quoted, and & → + (outside strings)
    sql = _convert_access_operators(sql)

    # 3. PlainText(expr) — custom Access VBA function; use expr directly
    sql = re.sub(r"\bPlainText\s*\(([^)]+)\)", r"\1", sql, flags=re.IGNORECASE)

    # 3. Trim$(expr) → TRIM(expr)
    sql = re.sub(r"\bTrim\$\s*\(", "TRIM(", sql, flags=re.IGNORECASE)

    # 4. Date() → CAST(GETDATE() AS DATE)
    sql = re.sub(r"\bDate\(\s*\)", "CAST(GETDATE() AS DATE)", sql, flags=re.IGNORECASE)

    # 5. DateValue(expr) → CAST(expr AS DATE)
    sql = re.sub(r"\bDateValue\s*\(([^)]+)\)", r"CAST(\1 AS DATE)", sql, flags=re.IGNORECASE)

    # 6. DateAdd("unit", n, expr) → DATEADD(unit, n, expr)
    def _replace_dateadd(m: re.Match) -> str:
        unit = m.group(1).strip("\"'").lower()
        tsql_unit = _ACCESS_DATEADD_MAP.get(unit, unit)
        return f"DATEADD({tsql_unit}, {m.group(2).strip()}, {m.group(3).strip()})"

    sql = re.sub(
        r"""\bDateAdd\s*\(\s*(["'][^"']+["'])\s*,\s*([^,]+?)\s*,\s*([^)]+?)\s*\)""",
        _replace_dateadd,
        sql,
        flags=re.IGNORECASE,
    )

    # 7. Format(expr, "fmt") — map Access/VBA format strings to T-SQL equivalents
    def _replace_format(m: re.Match) -> str:
        expr = m.group(1).strip()
        fmt = m.group(2).strip("\"'").lower()
        if fmt == "ww":
            return f"CAST(DATEPART(week, {expr}) AS VARCHAR)"
        if fmt == "q-yyyy":
            return f"(CAST(DATEPART(quarter, {expr}) AS VARCHAR) + '-' + FORMAT({expr}, 'yyyy'))"
        if fmt == "mmm-yyyy":
            return f"FORMAT({expr}, 'MMM-yyyy')"
        if fmt == "yyyy-mm":
            return f"FORMAT({expr}, 'yyyy-MM')"
        tsql_fmt = fmt.replace("mmmm", "MMMM").replace("mmm", "MMM").replace("mm", "MM")
        notes.append(f"Format() '{fmt}' → FORMAT(..., '{tsql_fmt}') — verify output")
        return f"FORMAT({expr}, '{tsql_fmt}')"

    sql = re.sub(
        r"""\bFormat\s*\(\s*([^,]+?)\s*,\s*(["'][^"']+["'])\s*\)""",
        _replace_format,
        sql,
        flags=re.IGNORECASE,
    )

    # 8. Remove ORDER BY when the view SELECT has no TOP (illegal in T-SQL views)
    sql = _strip_order_by_if_no_top(sql)

    return sql, notes


def _sort_views_topologically(views: list[dict]) -> list[dict]:
    """Return views sorted so each view is created after all views it references."""
    names = {v["name"] for v in views}
    by_name = {v["name"]: v for v in views}
    deps: dict[str, set[str]] = {}
    for v in views:
        sql_upper = v["sql"].upper()
        deps[v["name"]] = {
            n for n in names
            if n != v["name"] and re.search(r"\b" + re.escape(n.upper()) + r"\b", sql_upper)
        }
    in_degree = {n: len(d) for n, d in deps.items()}
    queue: deque[str] = deque(n for n, d in in_degree.items() if d == 0)
    ordered: list[str] = []
    while queue:
        node = queue.popleft()
        ordered.append(node)
        for n, d in deps.items():
            if node in d:
                d.discard(node)
                in_degree[n] -= 1
                if in_degree[n] == 0:
                    queue.append(n)
    for n in names:  # append any remaining (cyclic — unlikely for views)
        if n not in ordered:
            ordered.append(n)
    return [by_name[n] for n in ordered]


def get_dao_queries(db_path: Path) -> list[dict]:
    """Use DAO via win32com to extract saved query names, types, and SQL text."""
    if not HAS_WIN32COM:
        return []
    db = None
    try:
        dao_engine = win32com.client.Dispatch("DAO.DBEngine.120")
        db = dao_engine.OpenDatabase(str(db_path))
        queries = []
        for qdef in db.QueryDefs:
            name = qdef.Name
            if name.startswith("~") or name.startswith("MSys"):
                continue
            try:
                type_id = int(qdef.Type)
                sql_text = qdef.SQL or ""
            except Exception:
                continue
            queries.append(
                {
                    "name": name,
                    "sql": sql_text.strip(),
                    "type_id": type_id,
                    "type_name": DAO_QUERY_TYPE_NAMES.get(type_id, f"UNKNOWN({type_id})"),
                }
            )
        return queries
    except Exception as e:
        print(f"  [DAO warning] {e}")
        return []
    finally:
        if db is not None:
            try:
                db.Close()
            except Exception:
                pass


def get_linked_tables(conn, db_path: Path) -> list[dict]:
    """Detect linked tables (external data sources) in the Access database."""
    # Try ADOX first — gives us the datasource connection string
    if HAS_WIN32COM and os.getenv("ACCESS_USE_ADOX", "1") == "1":
        try:
            cat = win32com.client.Dispatch("ADOX.Catalog")
            cat.ActiveConnection = f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db_path};"
            linked = []
            for table in cat.Tables:
                if table.Name.startswith("MSys"):
                    continue
                if table.Type != "LINK":
                    continue
                entry: dict = {"name": table.Name}
                for prop_name in (
                    "Jet OLEDB:Link Datasource",
                    "Jet OLEDB:Link Provider String",
                    "Jet OLEDB:Remote Table Name",
                ):
                    try:
                        entry[prop_name] = table.Properties(prop_name).Value
                    except Exception:
                        pass
                linked.append(entry)
            return linked
        except Exception as e:
            print(f"  [linked tables ADOX warning] {e}")

    # Fallback: pyodbc — linked tables typically appear with table_type != "TABLE"
    linked = []
    cursor = conn.cursor()
    try:
        for row in cursor.tables():
            if row.table_name.startswith("MSys"):
                continue
            if row.table_type not in ("TABLE", "VIEW", "SYSTEM TABLE", ""):
                linked.append({"name": row.table_name, "type": row.table_type})
    except Exception:
        pass
    return linked


def sanitize_identifier(name: str) -> str:
    """Bracket-quote a T-SQL identifier, doubling any embedded ] so names
    containing ']', spaces, or reserved words can't break out of the brackets."""
    return "[" + str(name).replace("]", "]]") + "]"


def quote_literal(text: str) -> str:
    """Render a value as a Unicode T-SQL string literal (N'...'), doubling
    embedded single quotes. Use whenever an identifier/name is embedded in a
    string context such as OBJECT_ID(...) or a sys.* catalog name comparison."""
    return "N'" + str(text).replace("'", "''") + "'"


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
        col_lines.append(
            f"    CONSTRAINT {sanitize_identifier('PK_' + table_name)} PRIMARY KEY ({pk_cols})"
        )

    create_body = "\n".join(
        [
            f"CREATE TABLE {sanitize_identifier(table_name)} (",
            ",\n".join(col_lines),
            ");",
        ]
    )
    return (
        f"IF OBJECT_ID({quote_literal(sanitize_identifier(table_name))}, N'U') IS NULL\nBEGIN\n"
        + create_body
        + "\nEND"
    )


def build_insert_sql(table_name: str, col_descs, rows, has_identity: bool = False) -> list[str]:
    if not rows:
        return []

    tbl = sanitize_identifier(table_name)
    col_names = ", ".join(sanitize_identifier(d[0]) for d in col_descs)

    inserts = []
    for row in rows:
        values = ", ".join(escape_sql_string(v) for v in row)
        inserts.append(f"    ({values})")

    inner_lines: list[str] = []
    if has_identity:
        inner_lines.append(f"    SET IDENTITY_INSERT {tbl} ON;")

    batch_size = 1000
    for i in range(0, len(inserts), batch_size):
        batch = inserts[i : i + batch_size]
        sql = (
            f"    INSERT INTO {tbl} ({col_names})\n    VALUES\n"
            + ",\n".join(f"    {v}" for v in batch)
            + ";"
        )
        inner_lines.append(sql)

    if has_identity:
        inner_lines.append(f"    SET IDENTITY_INSERT {tbl} OFF;")

    block = (
        f"IF NOT EXISTS (SELECT 1 FROM {tbl})\nBEGIN\n"
        + "\n".join(inner_lines)
        + "\nEND"
    )
    return [block]


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


def resolve_db_paths(args: argparse.Namespace, access_dir: Path) -> list[Path]:
    if args.db_path:
        return [Path(args.db_path).resolve()]

    if args.db_name:
        return [(access_dir / args.db_name).resolve()]

    candidates = sorted(
        [
            p.resolve()
            for p in access_dir.iterdir()
            if p.is_file() and p.suffix.lower() in {".accdb", ".mdb"}
        ]
    )
    if len(candidates) == 1:
        print(f"Auto-selected database: {candidates[0].name}")
        return candidates

    if len(candidates) == 0:
        raise FileNotFoundError(
            f"No .accdb/.mdb file found in: {access_dir}\n"
            "Put your Access file there, or pass --db-name/--db-path explicitly."
        )

    print(f"Found {len(candidates)} Access files. Processing each into its own output folder.")
    for candidate in candidates:
        print(f"  - {candidate.name}")
    return candidates


def resolve_output_dir(repo_root: Path, output_dir_arg: str | None, db_path: Path, multi_db: bool) -> Path:
    if output_dir_arg:
        base_output_dir = Path(output_dir_arg).resolve()
        return (base_output_dir / db_path.stem).resolve() if multi_db else base_output_dir

    return (repo_root / "migration_output" / db_path.stem).resolve()


def extract_database(db_path: Path, output_dir: Path, access_dir: Path):
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

    print("Reading DAO queries...")
    dao_queries = get_dao_queries(db_path)
    if dao_queries:
        print(f"  Found {len(dao_queries)} saved queries")
    else:
        print("  No saved queries found (or DAO unavailable)")

    conn = get_connection(db_path)

    tables = get_tables(conn)
    print(f"Found {len(tables)} tables: {tables}")

    print("Checking for linked tables...")
    linked_tables = get_linked_tables(conn, db_path)
    if linked_tables:
        print(f"  *** WARNING: {len(linked_tables)} linked table(s) detected — external dependencies! ***")
        print("  These reference external data sources and will NOT be migrated automatically.")
        for lt in linked_tables:
            src = lt.get("Jet OLEDB:Link Datasource", "unknown source")
            print(f"    - {lt['name']} -> {src}")
        print()
    else:
        print("  No linked tables found.")

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
                f"IF NOT EXISTS (\n"
                f"    SELECT 1 FROM sys.foreign_keys\n"
                f"    WHERE name = {quote_literal(constraint_name)} AND parent_object_id = OBJECT_ID({quote_literal(sanitize_identifier(table_name))})\n"
                f")\n"
                f"    ALTER TABLE {sanitize_identifier(table_name)}\n"
                f"      ADD CONSTRAINT {sanitize_identifier(constraint_name)}\n"
                f"      FOREIGN KEY ({sanitize_identifier(fk['fk_column'])})\n"
                f"      REFERENCES {sanitize_identifier(fk['pk_table'])} ({sanitize_identifier(fk['pk_column'])});"
            )
            fk_statements.append(fk_sql)

        col_type_map = {col["name"]: access_type_to_tsql(col) for col in columns}
        col_nullable_map = {col["name"]: col.get("nullable", True) for col in columns}

        for idx in indexes:
            # Skip primary-key indexes — already created via CONSTRAINT in CREATE TABLE
            if idx.get("primary") or idx["name"] in (f"PK_{table_name}", "PrimaryKey"):
                continue
            if not idx["columns"]:
                continue
            # Skip indexes on LOB columns — SQL Server cannot use MAX/binary types as index keys
            lob_cols = [
                c for c in idx["columns"]
                if "MAX" in col_type_map.get(c, "").upper()
                or col_type_map.get(c, "").upper() == "XML"
            ]
            if lob_cols:
                index_statements.append(
                    f"-- SKIPPED index [{idx['name']}] on [{table_name}]: "
                    f"column(s) {lob_cols} have LOB/MAX type, not valid as index key in T-SQL"
                )
                continue
            unique_kw = "UNIQUE " if idx["unique"] else ""
            cols = ", ".join(sanitize_identifier(c) for c in idx["columns"])
            # Access allows multiple NULLs in a unique index; SQL Server does not.
            # Use a filtered index (WHERE col IS NOT NULL) to match Access semantics.
            where_clause = ""
            if idx["unique"]:
                nullable_idx_cols = [c for c in idx["columns"] if col_nullable_map.get(c, True)]
                if nullable_idx_cols:
                    conditions = " AND ".join(
                        f"{sanitize_identifier(c)} IS NOT NULL" for c in nullable_idx_cols
                    )
                    where_clause = f"\n  WHERE {conditions}"
            idx_sql = (
                f"IF NOT EXISTS (\n"
                f"    SELECT 1 FROM sys.indexes\n"
                f"    WHERE name = {quote_literal(idx['name'])} AND object_id = OBJECT_ID({quote_literal(sanitize_identifier(table_name))})\n"
                f")\n"
                f"    CREATE {unique_kw}INDEX {sanitize_identifier(idx['name'])}\n"
                f"      ON {sanitize_identifier(table_name)} ({cols}){where_clause};"
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

    view_queries = [q for q in dao_queries if q["type_id"] in VIEW_QUERY_TYPES]
    action_queries = [q for q in dao_queries if q["type_id"] not in VIEW_QUERY_TYPES]

    if view_queries:
        views_path = output_dir / "05_views.sql"
        sorted_views = _sort_views_topologically(view_queries)
        with open(views_path, "w", encoding="utf-8") as f:
            f.write(
                "-- Saved Queries exported as Views\n"
                "-- Source: Access saved SELECT/UNION queries\n"
                "-- Access SQL automatically converted to T-SQL.\n"
                "-- Views referencing Access form controls are commented out below.\n\n"
            )
            for q in sorted_views:
                # Skip views that reference Access runtime objects (form controls, subforms, reports, etc.)
                # Matches [Forms], [Form], [Parent], [Reports], [Report], [TempVars], [Me]
                if re.search(r"\[(forms?|parent|reports?|tempvars|me)\]", q["sql"], re.IGNORECASE):
                    f.write(
                        f"-- SKIPPED: {q['name']}\n"
                        "-- Reason: references Access form controls — cannot be a SQL view.\n"
                        "-- Convert manually to a stored procedure or parameterised query.\n"
                        "-- Original Access SQL:\n/*\n"
                    )
                    f.write(q["sql"])
                    f.write("\n*/\n\n")
                    continue
                converted_sql, conv_notes = access_sql_to_tsql(q["sql"])
                f.write(f"-- Query type: {q['type_name']}\n")
                for note in conv_notes:
                    f.write(f"-- Note: {note}\n")
                f.write(f"CREATE OR ALTER VIEW {sanitize_identifier(q['name'])} AS\n")
                f.write(converted_sql)
                if not converted_sql.endswith(";"):
                    f.write(";")
                f.write("\nGO\n\n")
        print(f"  Written: {views_path}")

    if action_queries:
        action_path = output_dir / "06_action_queries.sql"
        with open(action_path, "w", encoding="utf-8") as f:
            f.write(
                "-- Access Action Queries (reference only — not directly runnable as T-SQL)\n"
                "-- These are DELETE, UPDATE, APPEND, MAKE-TABLE, and other non-SELECT queries.\n"
                "-- Migrate manually as stored procedures or application logic as needed.\n\n"
            )
            for q in action_queries:
                f.write(f"-- ============================================================\n")
                f.write(f"-- Query: {q['name']}\n")
                f.write(f"-- Type:  {q['type_name']}\n")
                f.write(f"-- ============================================================\n")
                f.write("/*\n")
                f.write(q["sql"])
                f.write("\n*/\n\n")
        print(f"  Written: {action_path}")

    json_path = output_dir / "schema.json"

    def json_serializer(obj):
        if isinstance(obj, type):
            return obj.__name__
        raise TypeError(f"Object of type {type(obj)} is not JSON serializable")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                **schema_info,
                "_queries": [{"name": q["name"], "type_name": q["type_name"]} for q in dao_queries],
                "_linked_tables": linked_tables,
            },
            f,
            indent=2,
            default=json_serializer,
        )
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

        if linked_tables:
            f.write(f"## Linked Tables — External Dependencies ({len(linked_tables)})\n\n")
            f.write("> **WARNING:** These tables reference external data sources and will NOT be migrated automatically.\n\n")
            f.write("| Table Name | Data Source | Provider |\n")
            f.write("|------------|-------------|----------|\n")
            for lt in linked_tables:
                src = lt.get("Jet OLEDB:Link Datasource", "")
                provider = lt.get("Jet OLEDB:Link Provider String", "")
                f.write(f"| {lt['name']} | {src} | {provider} |\n")
            f.write("\n")

        if dao_queries:
            f.write(f"## Queries ({len(dao_queries)})\n\n")
            f.write("| Name | Type |\n")
            f.write("|------|------|\n")
            for q in dao_queries:
                f.write(f"| {q['name']} | {q['type_name']} |\n")
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


def main():
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    access_dir = Path(args.access_dir).resolve()
    access_dir.mkdir(parents=True, exist_ok=True)

    db_paths = resolve_db_paths(args, access_dir)
    multi_db = len(db_paths) > 1
    for db_path in db_paths:
        output_dir = resolve_output_dir(repo_root, args.output_dir, db_path, multi_db)
        extract_database(db_path, output_dir, access_dir)


if __name__ == "__main__":
    main()
