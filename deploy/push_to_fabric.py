"""
Deploy generated SQL migration scripts to Microsoft Fabric SQL Database using sqlcmd.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
from pathlib import Path


def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    default_output_dir = repo_root / "migration_output"

    parser = argparse.ArgumentParser(
        description=(
            "Execute generated SQL migration files against Microsoft Fabric SQL Database via sqlcmd."
        )
    )
    parser.add_argument("--server", required=True, help="Fabric SQL endpoint host name")
    parser.add_argument("--database", required=True, help="Target Fabric SQL database name")
    parser.add_argument(
        "--db-name",
        default=None,
        help=(
            "Folder name under migration_output to deploy from. "
            "If omitted, auto-selects when exactly one folder exists."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=str(default_output_dir),
        help="Base folder containing migration output subfolders (default: ./migration_output)",
    )
    parser.add_argument(
        "--auth",
        choices=["interactive", "default", "integrated", "password"],
        default="interactive",
        help=(
            "Authentication mode: interactive/default/integrated/password. "
            "Default is interactive (passwordless Entra login)."
        ),
    )
    parser.add_argument(
        "--username",
        default=None,
        help="Username for password auth only (required with --auth password)",
    )
    parser.add_argument(
        "--include-views",
        action="store_true",
        help="Run 05_views.sql if present (Access SQL may need adaptation)",
    )
    parser.add_argument(
        "--skip-fk",
        action="store_true",
        help="Skip 02_foreign_keys.sql even if present",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print commands without executing sqlcmd",
    )
    return parser.parse_args()


def detect_sqlcmd_binary() -> str:
    sqlcmd_path = shutil.which("sqlcmd")
    if not sqlcmd_path:
        raise RuntimeError(
            "sqlcmd not found in PATH. Install it and rerun. Example check: sqlcmd -?"
        )
    return sqlcmd_path


def is_go_sqlcmd(sqlcmd_path: str) -> bool:
    try:
        result = subprocess.run(
            [sqlcmd_path, "--help"],
            capture_output=True,
            text=True,
            check=False,
        )
        help_text = (result.stdout or "") + "\n" + (result.stderr or "")
        return "--authentication-method" in help_text
    except Exception:
        return False


def resolve_migration_dir(base_output_dir: Path, db_name: str | None) -> Path:
    if db_name:
        target = (base_output_dir / db_name).resolve()
        if not target.exists() or not target.is_dir():
            raise FileNotFoundError(f"Migration folder not found: {target}")
        return target

    candidates = sorted([p for p in base_output_dir.iterdir() if p.is_dir()])
    if len(candidates) == 1:
        print(f"Auto-selected migration folder: {candidates[0].name}")
        return candidates[0].resolve()

    if not candidates:
        raise FileNotFoundError(f"No database folders found under: {base_output_dir}")

    names = ", ".join(p.name for p in candidates)
    raise RuntimeError(
        "Multiple migration folders found. Pass --db-name explicitly. "
        f"Available: {names}"
    )


def build_auth_args(args: argparse.Namespace, use_go_sqlcmd: bool) -> list[str]:
    if args.auth == "integrated":
        return ["-E"]

    if args.auth == "password":
        if not args.username:
            raise ValueError("--username is required when --auth password")
        password = os.getenv("FABRIC_SQL_PASSWORD")
        if not password:
            raise ValueError(
                "FABRIC_SQL_PASSWORD is not set. Set env var or use --auth interactive/default."
            )
        return ["-U", args.username, "-P", password]

    if args.auth == "default":
        if use_go_sqlcmd:
            return ["--authentication-method", "ActiveDirectoryDefault"]
        # Legacy sqlcmd does not support ActiveDirectoryDefault.
        # Fallback to interactive token flow without password.
        return ["-G"]

    # interactive (default)
    return ["-G"]


def get_script_plan(migration_dir: Path, include_views: bool, skip_fk: bool) -> list[Path]:
    mandatory = migration_dir / "01_create_tables.sql"
    if not mandatory.exists():
        raise FileNotFoundError(f"Required file missing: {mandatory}")

    ordered_names = ["01_create_tables.sql", "04_insert_data.sql", "03_indexes.sql"]
    if not skip_fk:
        ordered_names.append("02_foreign_keys.sql")
    if include_views:
        ordered_names.append("05_views.sql")

    plan: list[Path] = []
    for name in ordered_names:
        path = migration_dir / name
        if path.exists():
            plan.append(path)
    return plan


def format_command_for_log(parts: list[str], hide_password: bool) -> str:
    if not hide_password:
        return " ".join(parts)

    redacted: list[str] = []
    i = 0
    while i < len(parts):
        token = parts[i]
        if token == "-P" and i + 1 < len(parts):
            redacted.extend(["-P", "***"])
            i += 2
            continue
        redacted.append(token)
        i += 1
    return " ".join(redacted)


def run_plan(
    sqlcmd_path: str,
    server: str,
    database: str,
    auth_args: list[str],
    scripts: list[Path],
    dry_run: bool,
) -> None:
    base_cmd = [sqlcmd_path, "-S", server, "-d", database, "-b", "-I"] + auth_args
    hide_password = "-P" in auth_args

    if dry_run:
        print("DRY RUN: sqlcmd commands to execute")

    for script in scripts:
        cmd = base_cmd + ["-i", str(script)]
        print(f"\n==> {script.name}")
        print(format_command_for_log(cmd, hide_password=hide_password))

        if dry_run:
            continue

        result = subprocess.run(cmd, check=False)
        if result.returncode != 0:
            raise RuntimeError(
                f"sqlcmd failed for {script.name} with exit code {result.returncode}"
            )


def main() -> None:
    args = parse_args()
    base_output_dir = Path(args.output_dir).resolve()

    if not base_output_dir.exists() or not base_output_dir.is_dir():
        raise FileNotFoundError(f"Output directory not found: {base_output_dir}")

    sqlcmd_path = detect_sqlcmd_binary()
    use_go_sqlcmd = is_go_sqlcmd(sqlcmd_path)

    migration_dir = resolve_migration_dir(base_output_dir, args.db_name)
    scripts = get_script_plan(
        migration_dir=migration_dir,
        include_views=args.include_views,
        skip_fk=args.skip_fk,
    )
    auth_args = build_auth_args(args, use_go_sqlcmd)

    print(f"Migration folder: {migration_dir}")
    print(f"Scripts to run ({len(scripts)}):")
    for script in scripts:
        print(f"  - {script.name}")

    if args.auth == "default" and not use_go_sqlcmd:
        print(
            "Note: legacy sqlcmd detected. --auth default falls back to interactive (-G)."
        )

    run_plan(
        sqlcmd_path=sqlcmd_path,
        server=args.server,
        database=args.database,
        auth_args=auth_args,
        scripts=scripts,
        dry_run=args.dry_run,
    )

    if args.dry_run:
        print("\nDry run complete.")
    else:
        print("\nDeployment complete.")


if __name__ == "__main__":
    main()
