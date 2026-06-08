"""Tests for the pure helpers in deploy/push_to_fabric.py."""

from argparse import Namespace

import pytest

import push_to_fabric as p


# --------------------------------------------------------------------------- #
# connection string parsing
# --------------------------------------------------------------------------- #


def test_parse_real_fabric_connection_string():
    cs = ("Data Source=abc-def.database.fabric.microsoft.com,1433;"
          "Initial Catalog=new_northwind-c1d30aea;Encrypt=True;"
          "Authentication=Active Directory Interactive")
    server, database = p.parse_connection_string(cs)
    assert server == "abc-def.database.fabric.microsoft.com,1433"
    assert database == "new_northwind-c1d30aea"


def test_parse_connection_string_server_database_aliases():
    assert p.parse_connection_string("Server=s;Database=d") == ("s", "d")


def test_parse_connection_string_missing_parts_raise():
    with pytest.raises(ValueError):
        p.parse_connection_string("Initial Catalog=d")      # no server
    with pytest.raises(ValueError):
        p.parse_connection_string("Data Source=s")          # no database


# --------------------------------------------------------------------------- #
# script plan ordering
# --------------------------------------------------------------------------- #


def _make_scripts(tmp_path, names):
    for n in names:
        (tmp_path / n).write_text("-- sql", encoding="utf-8")


def test_script_plan_default_order(tmp_path):
    _make_scripts(tmp_path, ["01_create_tables.sql", "02_foreign_keys.sql",
                             "03_indexes.sql", "04_insert_data.sql", "05_views.sql"])
    plan = [pth.name for pth in p.get_script_plan(tmp_path, include_views=False, skip_fk=False)]
    assert plan == ["01_create_tables.sql", "04_insert_data.sql",
                    "03_indexes.sql", "02_foreign_keys.sql"]   # views excluded by default


def test_script_plan_include_views_and_skip_fk(tmp_path):
    _make_scripts(tmp_path, ["01_create_tables.sql", "02_foreign_keys.sql",
                             "03_indexes.sql", "04_insert_data.sql", "05_views.sql"])
    plan = [pth.name for pth in p.get_script_plan(tmp_path, include_views=True, skip_fk=True)]
    assert plan == ["01_create_tables.sql", "04_insert_data.sql",
                    "03_indexes.sql", "05_views.sql"]          # no FK, views appended last


def test_script_plan_missing_mandatory_raises(tmp_path):
    _make_scripts(tmp_path, ["04_insert_data.sql"])
    with pytest.raises(FileNotFoundError):
        p.get_script_plan(tmp_path, include_views=False, skip_fk=False)


def test_script_plan_skips_absent_optional_files(tmp_path):
    _make_scripts(tmp_path, ["01_create_tables.sql", "04_insert_data.sql"])
    plan = [pth.name for pth in p.get_script_plan(tmp_path, include_views=True, skip_fk=False)]
    assert plan == ["01_create_tables.sql", "04_insert_data.sql"]


# --------------------------------------------------------------------------- #
# auth args
# --------------------------------------------------------------------------- #


def test_build_auth_args_modes():
    assert p.build_auth_args(Namespace(auth="integrated", username=None), False) == ["-E"]
    assert p.build_auth_args(Namespace(auth="interactive", username=None), False) == ["-G"]
    assert p.build_auth_args(Namespace(auth="default", username=None), True) == \
        ["--authentication-method", "ActiveDirectoryDefault"]
    assert p.build_auth_args(Namespace(auth="default", username=None), False) == ["-G"]  # legacy fallback


def test_build_auth_args_password_requires_username_and_env(monkeypatch):
    monkeypatch.delenv("FABRIC_SQL_PASSWORD", raising=False)
    with pytest.raises(ValueError):
        p.build_auth_args(Namespace(auth="password", username=None), False)  # no username
    with pytest.raises(ValueError):
        p.build_auth_args(Namespace(auth="password", username="u"), False)   # no env password
    monkeypatch.setenv("FABRIC_SQL_PASSWORD", "s3cret")
    assert p.build_auth_args(Namespace(auth="password", username="u"), False) == \
        ["-U", "u", "-P", "s3cret"]


# --------------------------------------------------------------------------- #
# log redaction
# --------------------------------------------------------------------------- #


def test_format_command_redacts_password():
    parts = ["sqlcmd", "-S", "srv", "-U", "u", "-P", "s3cret", "-i", "x.sql"]
    assert p.format_command_for_log(parts, hide_password=True) == \
        "sqlcmd -S srv -U u -P *** -i x.sql"
    assert "s3cret" in p.format_command_for_log(parts, hide_password=False)
