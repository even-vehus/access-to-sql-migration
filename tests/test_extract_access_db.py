"""Tests for the SQL-generation core of migration/extract_access_db.py."""

from datetime import date, datetime
from decimal import Decimal

import pytest

import extract_access_db as m


def _col(name, type_name="NVARCHAR", *, size=None, nullable=True, type_code=None,
         precision=None, scale=0):
    """Build a column dict shaped like get_column_info() output."""
    return {
        "name": name,
        "type_name": type_name,
        "column_size": size,
        "decimal_digits": scale,
        "nullable": nullable,
        "column_def": None,
        "ordinal_position": 1,
        "_type_code": type_code,
        "_precision": precision,
    }


# --------------------------------------------------------------------------- #
# #5 — identifier / literal escaping
# --------------------------------------------------------------------------- #


def test_sanitize_identifier_brackets_and_doubles_close_bracket():
    assert m.sanitize_identifier("Orders") == "[Orders]"
    assert m.sanitize_identifier("Order]Id") == "[Order]]Id]"
    # An apostrophe is fine *inside* brackets and must NOT be doubled there.
    assert m.sanitize_identifier("Order's") == "[Order's]"


def test_quote_literal_doubles_single_quote():
    assert m.quote_literal("O'Brien") == "N'O''Brien'"
    assert m.quote_literal("[Foo's]") == "N'[Foo''s]'"


def test_create_table_escapes_hostile_table_and_column_names():
    cols = [_col("Order]Id", "LONG", nullable=False, type_code=int),
            _col("Cust's Name", "NVARCHAR", size=50, type_code=str)]
    ddl = m.build_create_table_sql("O'Brien]Orders", cols, ["Order]Id"], None)
    # OBJECT_ID literal: ] doubled for the identifier, ' doubled for the literal.
    assert "OBJECT_ID(N'[O''Brien]]Orders]', N'U')" in ddl
    assert "CREATE TABLE [O'Brien]]Orders] (" in ddl
    assert "[Order]]Id] INT NOT NULL" in ddl
    assert "[Cust's Name] NVARCHAR(50)" in ddl
    assert "CONSTRAINT [PK_O'Brien]]Orders] PRIMARY KEY ([Order]]Id])" in ddl


# --------------------------------------------------------------------------- #
# escape_sql_string — data literals
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize("value,expected", [
    (None, "NULL"),
    (True, "1"),
    (False, "0"),
    (42, "42"),
    (3.5, "3.5"),
    (Decimal("1.5"), "1.5"),
    (date(2021, 6, 8), "'2021-06-08'"),
    (datetime(2021, 6, 8, 9, 30, 15, 123456), "'2021-06-08 09:30:15.123'"),
    (b"\xde\xad", "0xdead"),
    ("O'Brien", "N'O''Brien'"),
    ("plain", "N'plain'"),
])
def test_escape_sql_string(value, expected):
    assert m.escape_sql_string(value) == expected


# --------------------------------------------------------------------------- #
# access_type_to_tsql — type mapping
# --------------------------------------------------------------------------- #


def test_type_mapping_by_python_type_code():
    assert m.access_type_to_tsql(_col("c", type_code=str, size=50)) == "NVARCHAR(50)"
    assert m.access_type_to_tsql(_col("c", type_code=str, size=0)) == "NVARCHAR(MAX)"
    assert m.access_type_to_tsql(_col("c", type_code=str, size=8000)) == "NVARCHAR(MAX)"
    assert m.access_type_to_tsql(_col("c", type_code=int)) == "INT"
    assert m.access_type_to_tsql(_col("c", type_code=bool)) == "BIT"
    assert m.access_type_to_tsql(_col("c", type_code=datetime)) == "DATETIME2"
    assert m.access_type_to_tsql(_col("c", type_code=Decimal, precision=10, scale=2)) == "DECIMAL(10,2)"


def test_type_mapping_by_access_type_name():
    assert m.access_type_to_tsql(_col("c", "COUNTER")) == "INT IDENTITY(1,1)"
    assert m.access_type_to_tsql(_col("c", "MEMO")) == "NVARCHAR(MAX)"
    assert m.access_type_to_tsql(_col("c", "CURRENCY")) == "DECIMAL(19,4)"
    assert m.access_type_to_tsql(_col("c", "YESNO")) == "BIT"
    assert m.access_type_to_tsql(_col("c", "GUID")) == "UNIQUEIDENTIFIER"


def test_unknown_type_falls_back_with_comment():
    out = m.access_type_to_tsql(_col("c", "WEIRDTYPE"))
    assert out.startswith("NVARCHAR(255)")
    assert "WEIRDTYPE" in out


# --------------------------------------------------------------------------- #
# access_sql_to_tsql — best-effort conversion
# --------------------------------------------------------------------------- #


def test_concatenation_and_string_literals():
    out, _ = m.access_sql_to_tsql('"x" & "y"')
    # & -> + and double-quoted literals -> single-quoted (exact spacing is cosmetic)
    assert " ".join(out.split()) == "'x' + 'y'"


def test_bang_field_separator_but_preserve_not_equal():
    out, _ = m.access_sql_to_tsql("SELECT t!f FROM t WHERE a != b")
    assert "t.f" in out
    assert "a != b" in out


def test_date_and_trim_and_dateadd():
    assert m.access_sql_to_tsql("Date()")[0] == "CAST(GETDATE() AS DATE)"
    assert m.access_sql_to_tsql("Trim$(x)")[0] == "TRIM(x)"
    assert m.access_sql_to_tsql('DateAdd("d", 1, x)')[0] == "DATEADD(day, 1, x)"


def test_combined_conversion():
    out, _ = m.access_sql_to_tsql('Trim$([t]![f]) & "x"')
    assert " ".join(out.split()) == "TRIM([t].[f]) + 'x'"


def test_strip_order_by_without_top_but_keep_with_top():
    assert m._strip_order_by_if_no_top("SELECT a FROM t ORDER BY a") == "SELECT a FROM t"
    kept = m._strip_order_by_if_no_top("SELECT TOP 5 a FROM t ORDER BY a")
    assert "ORDER BY" in kept


# --------------------------------------------------------------------------- #
# view dependency ordering
# --------------------------------------------------------------------------- #


def test_views_sorted_so_dependencies_come_first():
    views = [
        {"name": "vDependent", "sql": "SELECT * FROM vBase"},
        {"name": "vBase", "sql": "SELECT * FROM SomeTable"},
    ]
    ordered = [v["name"] for v in m._sort_views_topologically(views)]
    assert ordered.index("vBase") < ordered.index("vDependent")


# --------------------------------------------------------------------------- #
# build_insert_sql
# --------------------------------------------------------------------------- #


def test_insert_sql_empty_rows_returns_nothing():
    assert m.build_insert_sql("T", [("id",)], []) == []


def test_insert_sql_with_identity_and_escaping():
    blocks = m.build_insert_sql("T", [("id",), ("name",)], [(1, "a'b")], has_identity=True)
    sql = "\n".join(blocks)
    assert "IF NOT EXISTS (SELECT 1 FROM [T])" in sql
    assert "SET IDENTITY_INSERT [T] ON;" in sql
    assert "SET IDENTITY_INSERT [T] OFF;" in sql
    assert "INSERT INTO [T] ([id], [name])" in sql
    assert "(1, N'a''b')" in sql
