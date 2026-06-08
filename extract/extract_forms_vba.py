"""
Access Forms, VBA, Reports & Macros Extractor
Uses Access.Application COM automation to extract UI components and code
that are not accessible via ODBC/ADOX.

Outputs forms_vba.json and forms_vba_summary.md per database.
"""

from __future__ import annotations

import argparse
import atexit
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path

import win32com.client
import win32process

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CONTROL_TYPE_NAMES: dict[int, str] = {
    100: "Label",
    104: "CommandButton",
    105: "OptionButton",
    106: "CheckBox",
    109: "TextBox",
    110: "ListBox",
    111: "ComboBox",
    112: "Subform",
    114: "ObjectFrame",
    118: "PageBreak",
    122: "ToggleButton",
    124: "TabControl",
    126: "Page",
    128: "CustomControl",
    130: "BoundObjectFrame",
    134: "Attachment",
}

# Access object type constants for SaveAsText / DoCmd
# (correspond to Access VBA enum values: acForm=2, acReport=3, acMacro=4, acModule=5)
AC_FORM = 2
AC_REPORT = 3
AC_MACRO = 4
AC_MODULE = 5
# acViewDesign=1 (NOT 2 — that is acViewPreview)
AC_VIEW_DESIGN = 1
# acSaveNo=2 (NOT 0 — that is acSavePrompt, which shows a dialog)
AC_SAVE_NO = 2

# VBA component types
VB_COMPONENT_TYPES: dict[int, str] = {
    1: "StandardModule",
    2: "ClassModule",
    3: "MSForm",
    100: "Document",
}

# Event properties to scan on controls
EVENT_PROPERTIES = [
    "OnClick",
    "OnDblClick",
    "BeforeUpdate",
    "AfterUpdate",
    "OnChange",
    "OnEnter",
    "OnExit",
    "OnGotFocus",
    "OnLostFocus",
    "OnOpen",
    "OnClose",
    "OnLoad",
    "OnCurrent",
    "OnDirty",
    "OnDelete",
    "BeforeInsert",
    "AfterInsert",
]

# Patterns for scanning VBA source
RE_API_DECLARATION = re.compile(
    r"^\s*(?:Public|Private)?\s*Declare\s+(?:Function|Sub)\s+",
    re.MULTILINE | re.IGNORECASE,
)
RE_EXTERNAL_REF = re.compile(
    r"\b(CreateObject|GetObject|Shell)\b", re.IGNORECASE
)

# Global references for cleanup / scoped process termination
_access_app = None
_access_pid: int | None = None

# Timeout for individual COM operations (seconds)
COM_TIMEOUT = 30


class COMTimeoutError(Exception):
    """Raised when a COM operation exceeds the allowed timeout."""


def _run_with_timeout(func, timeout: int = COM_TIMEOUT):
    """Run a callable in a thread; if it exceeds timeout, kill MSACCESS and raise."""
    result = [None]
    error = [None]

    def _worker():
        try:
            result[0] = func()
        except Exception as exc:
            error[0] = exc

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    t.join(timeout)

    if t.is_alive():
        # The COM call is hung (Access likely raised a modal dialog).
        # Terminate ONLY the instance we launched — never a broad image kill,
        # which would destroy other Access windows the user has open with
        # unsaved work. /T also reaps any child processes of that PID.
        if _access_pid:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(_access_pid)],
                    capture_output=True,
                )
            except Exception:
                pass
        else:
            print(
                "  [warning] Access COM call hung but its PID is unknown; "
                "skipping process kill to avoid terminating unrelated Access "
                "instances. An orphaned MSACCESS.EXE may remain - close it manually."
            )
        raise COMTimeoutError(
            f"COM operation timed out after {timeout}s (Access likely showed a modal dialog)"
        )

    if error[0] is not None:
        raise error[0]
    return result[0]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    default_access_dir = repo_root / "access_databases"

    parser = argparse.ArgumentParser(
        description="Extract forms, reports, VBA modules, and macros from Access databases via COM automation."
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
            "When multiple databases are processed, each gets a subfolder under this path."
        ),
    )
    parser.add_argument(
        "--password",
        default=None,
        help="Password for opening a password-protected database.",
    )
    return parser.parse_args()


# ---------------------------------------------------------------------------
# Path resolution (duplicated from extract_access_db.py for independence)
# ---------------------------------------------------------------------------


def resolve_db_paths(args: argparse.Namespace, access_dir: Path) -> list[Path]:
    if args.db_path:
        return [Path(args.db_path).resolve()]

    if args.db_name:
        return [(access_dir / args.db_name).resolve()]

    candidates = sorted(
        p.resolve()
        for p in access_dir.iterdir()
        if p.is_file() and p.suffix.lower() in {".accdb", ".mdb"}
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


def resolve_output_dir(
    repo_root: Path, output_dir_arg: str | None, db_path: Path, multi_db: bool
) -> Path:
    if output_dir_arg:
        base = Path(output_dir_arg).resolve()
        return (base / db_path.stem).resolve() if multi_db else base
    return (repo_root / "migration_output" / db_path.stem).resolve()


# ---------------------------------------------------------------------------
# COM lifecycle helpers
# ---------------------------------------------------------------------------


def _cleanup_access_app():
    """atexit handler — kill orphaned Access process if script crashes."""
    global _access_app, _access_pid
    if _access_app is not None:
        try:
            _access_app.Quit()
        except Exception:
            pass
        _access_app = None
    _access_pid = None


atexit.register(_cleanup_access_app)


def _get_access_pid(app) -> int | None:
    """Resolve the OS process ID of the Access instance behind this COM app.

    Uses the main Access window handle so the timeout watchdog can terminate
    only the instance we launched — never other Access windows the user may
    have open. Returns None if it can't be determined.
    """
    try:
        hwnd = int(app.hWndAccessApp)
    except Exception:
        return None
    if not hwnd:
        return None
    try:
        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid) or None
    except Exception:
        return None


def open_access_app(db_path: Path, password: str | None = None):
    """Open an Access database via COM and return the Application object."""
    global _access_app, _access_pid
    app = win32com.client.Dispatch("Access.Application")
    # Assign to global only after object creation, before DB open.
    # The atexit handler will call Quit() if the script crashes here.
    _access_app = app

    # Prevent AutoExec macros from running (set before opening DB)
    try:
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    except Exception:
        pass

    if password:
        app.OpenCurrentDatabase(str(db_path), False, password)
    else:
        app.OpenCurrentDatabase(str(db_path))

    # Set invisible after opening — some Access versions reject this before a DB is open
    try:
        app.Visible = False
    except Exception:
        pass

    # Capture the PID now, while Access is healthy, so the timeout watchdog can
    # later kill only this instance (not every MSACCESS.EXE on the machine).
    _access_pid = _get_access_pid(app)

    return app


def close_access_app(app) -> None:
    """Safely close the Access COM application."""
    global _access_app, _access_pid
    try:
        app.CloseCurrentDatabase()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass
    _access_app = None
    _access_pid = None


# ---------------------------------------------------------------------------
# Form extraction
# ---------------------------------------------------------------------------


def _safe_getattr(obj, attr: str, default=None):
    """Read a COM property that may not exist or may error."""
    try:
        val = getattr(obj, attr)
        # COM may return empty variant as None or empty string
        if val is None or val == "":
            return default
        return val
    except Exception:
        return default


def _get_control_events(control) -> list[str]:
    """Return list of event names that have [Event Procedure] handlers."""
    events: list[str] = []
    for prop_name in EVENT_PROPERTIES:
        try:
            val = getattr(control, prop_name)
            if val and "[Event Procedure]" in str(val):
                events.append(prop_name)
        except Exception:
            continue
    return events


def _extract_control(control) -> dict:
    """Extract properties from a single form/report control."""
    ctrl_type_int = _safe_getattr(control, "ControlType", -1)
    ctrl_type_name = CONTROL_TYPE_NAMES.get(ctrl_type_int, f"Unknown({ctrl_type_int})")

    info: dict = {
        "name": _safe_getattr(control, "Name", ""),
        "type": ctrl_type_name,
        "type_id": ctrl_type_int,
    }

    # Common data-binding properties
    info["control_source"] = _safe_getattr(control, "ControlSource")
    info["default_value"] = _safe_getattr(control, "DefaultValue")
    info["validation_rule"] = _safe_getattr(control, "ValidationRule")
    info["validation_text"] = _safe_getattr(control, "ValidationText")
    info["visible"] = _safe_getattr(control, "Visible", True)
    info["enabled"] = _safe_getattr(control, "Enabled", True)
    info["locked"] = _safe_getattr(control, "Locked", False)

    # Combo/List box specifics
    if ctrl_type_int in (110, 111):  # ListBox, ComboBox
        info["row_source"] = _safe_getattr(control, "RowSource")
        info["row_source_type"] = _safe_getattr(control, "RowSourceType")

    # Subform specifics
    if ctrl_type_int == 112:  # Subform
        info["source_object"] = _safe_getattr(control, "SourceObject")
        info["link_child_fields"] = _safe_getattr(control, "LinkChildFields")
        info["link_master_fields"] = _safe_getattr(control, "LinkMasterFields")

    # Events
    info["events"] = _get_control_events(control)

    return info


# ---------------------------------------------------------------------------
# SaveAsText parser for forms/reports
# ---------------------------------------------------------------------------

# Regex patterns for parsing SaveAsText output
_RE_BEGIN_CONTROL = re.compile(r"^\s+Begin\s+(\w+)\s*$")
_RE_END = re.compile(r"^\s+End\s*$")
# Structural Begin (with or without a type name) — NOT a binary-data "PropName = Begin" line
_RE_BEGIN_STRUCT = re.compile(r"^\s+Begin(\s|$)")
# Binary data block start:  PropName = Begin  (hex data follows, terminated by End)
_RE_BEGIN_BINARY = re.compile(r"^\s+\w+\s*=\s*Begin\s*$")

# Map SaveAsText control type keywords to our type names
_TEXT_CONTROL_TYPES: dict[str, tuple[str, int]] = {
    "TextBox": ("TextBox", 109),
    "ComboBox": ("ComboBox", 111),
    "ListBox": ("ListBox", 110),
    "CommandButton": ("CommandButton", 104),
    "Label": ("Label", 100),
    "CheckBox": ("CheckBox", 106),
    "OptionButton": ("OptionButton", 105),
    "Subform": ("Subform", 112),
    "SubForm": ("Subform", 112),
    "ToggleButton": ("ToggleButton", 122),
    "TabControl": ("TabControl", 124),
    "Page": ("Page", 126),
    "ObjectFrame": ("ObjectFrame", 114),
    "BoundObjectFrame": ("BoundObjectFrame", 130),
    "CustomControl": ("CustomControl", 128),
    "PageBreak": ("PageBreak", 118),
    "Attachment": ("Attachment", 134),
    "OptionGroup": ("OptionGroup", 107),
    "Image": ("Image", 103),
    "Line": ("Line", 102),
    "Rectangle": ("Rectangle", 101),
}

# Event properties recognized in SaveAsText output
_TEXT_EVENT_PROPS = {
    "OnClick", "OnDblClick", "BeforeUpdate", "AfterUpdate",
    "OnChange", "OnEnter", "OnExit", "OnGotFocus", "OnLostFocus",
    "OnOpen", "OnClose", "OnLoad", "OnCurrent", "OnDirty",
    "OnDelete", "BeforeInsert", "AfterInsert",
}


def _read_access_text_file(path: Path) -> str:
    """Read a file produced by Access SaveAsText, handling UTF-16 LE/BE BOM."""
    raw = path.read_bytes()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", errors="replace")
    return raw.decode("utf-8", errors="replace")


def _cleanup_file(path: Path) -> None:
    """Remove a temp file, ignoring errors."""
    try:
        path.unlink(missing_ok=True)
    except Exception:
        pass


def _parse_form_text(text: str, name: str, is_report: bool = False) -> dict | None:
    """Parse Access SaveAsText output into structured form/report data.

    The format is a nested Begin/End block structure with property assignments.
    """
    if not text.strip():
        return None

    record_source: str | None = None
    controls: list[dict] = []
    group_levels: list[dict] = []

    # Extract top-level RecordSource
    for line in text.splitlines():
        m = re.match(r'^\s+RecordSource\s*=\s*"(.+?)"\s*$', line)
        if m:
            record_source = m.group(1)
            break
        # Also try without quotes (for simple table names)
        m = re.match(r"^\s+RecordSource\s*=\s*(.+?)\s*$", line)
        if m and not m.group(1).startswith('"'):
            record_source = m.group(1)
            break

    # Parse control blocks
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        m = _RE_BEGIN_CONTROL.match(line)
        if m:
            ctrl_type_keyword = m.group(1)
            if ctrl_type_keyword in _TEXT_CONTROL_TYPES:
                ctrl_data, end_i = _parse_control_block(lines, i, ctrl_type_keyword)
                if ctrl_data:
                    controls.append(ctrl_data)
                i = end_i
            else:
                i += 1
        else:
            # Check for GroupLevel in reports
            if is_report and "GroupLevel" in line:
                gl = _parse_group_level_line(line)
                if gl:
                    group_levels.append(gl)
            i += 1

    result: dict = {
        "name": name,
        "record_source": record_source,
        "controls": controls,
    }
    if is_report:
        result["group_levels"] = group_levels

    return result


def _parse_control_block(lines: list[str], start: int, ctrl_type_keyword: str) -> tuple[dict | None, int]:
    """Parse a single control's Begin...End block from SaveAsText output."""
    type_name, type_id = _TEXT_CONTROL_TYPES[ctrl_type_keyword]
    props: dict[str, str] = {}
    depth = 1
    i = start + 1

    binary_depth = 0  # tracks nested "PropName = Begin ... End" binary data blocks

    while i < len(lines) and depth > 0:
        line = lines[i]
        if _RE_BEGIN_BINARY.match(line):
            # Binary data block (e.g. GUID = Begin / PrtMip = Begin) — don't touch structural depth
            binary_depth += 1
        elif binary_depth > 0 and _RE_END.match(line):
            binary_depth -= 1
        elif _RE_BEGIN_STRUCT.match(line):
            # Structural Begin (with or without a type word, e.g. "Begin Label" or bare "Begin")
            depth += 1
        elif _RE_END.match(line):
            depth -= 1
            if depth == 0:
                break
        elif depth == 1 and binary_depth == 0:
            # Only read properties at the first nesting level of this control
            m = re.match(r'\s+(\w+)\s*=\s*"(.*?)"\s*$', line)
            if m:
                props[m.group(1)] = m.group(2)
            else:
                m = re.match(r"\s+(\w+)\s*=\s*(.+?)\s*$", line)
                if m:
                    props[m.group(1)] = m.group(2)
        i += 1

    ctrl_name = props.get("Name", "")
    if not ctrl_name:
        return None, i

    info: dict = {
        "name": ctrl_name,
        "type": type_name,
        "type_id": type_id,
        "control_source": props.get("ControlSource"),
        "default_value": props.get("DefaultValue"),
        "validation_rule": props.get("ValidationRule"),
        "validation_text": props.get("ValidationText"),
        "visible": props.get("Visible", "True") != "0",
        "enabled": props.get("Enabled", "True") != "0",
        "locked": props.get("Locked", "False") == "-1",
    }

    # Combo/List box specifics
    if type_id in (110, 111):
        info["row_source"] = props.get("RowSource")
        info["row_source_type"] = props.get("RowSourceType")

    # Subform specifics
    if type_id == 112:
        info["source_object"] = props.get("SourceObject")
        info["link_child_fields"] = props.get("LinkChildFields")
        info["link_master_fields"] = props.get("LinkMasterFields")

    # Events
    events: list[str] = []
    for prop_name in _TEXT_EVENT_PROPS:
        val = props.get(prop_name, "")
        if "[Event Procedure]" in val:
            events.append(prop_name)
    info["events"] = events

    return info, i


def _parse_group_level_line(line: str) -> dict | None:
    """Parse a GroupLevel definition from report text."""
    m = re.search(r"GroupLevel\s*=\s*(.+)", line)
    if m:
        return {"field": m.group(1).strip(), "group_on": 0, "sort_order": 0}
    return None


def extract_forms(app, warnings: list[str]) -> list[dict]:
    """Extract all forms from the open Access database.

    Uses a two-pass approach:
    1. Try SaveAsText (reliable, no design-view dialogs) to get form text, then parse.
    2. Fall back to opening in design view for any forms where SaveAsText fails.
    """
    forms: list[dict] = []

    try:
        all_forms = app.CurrentProject.AllForms
    except Exception as exc:
        warnings.append(f"Could not enumerate forms: {exc}")
        return forms

    form_count = all_forms.Count
    print(f"  Found {form_count} form(s)")

    tmp_dir = Path(tempfile.mkdtemp(prefix="access_forms_"))

    for i in range(form_count):
        form_obj = all_forms.Item(i)
        form_name = form_obj.Name
        print(f"    Processing form: {form_name}")

        # --- Primary method: SaveAsText ---
        out_file = tmp_dir / f"{form_name}.txt"
        try:
            app.SaveAsText(AC_FORM, form_name, str(out_file))
            form_text = _read_access_text_file(out_file)
            parsed = _parse_form_text(form_text, form_name)
            if parsed:
                forms.append(parsed)
                _cleanup_file(out_file)
                continue
        except Exception:
            pass

        # --- Fallback: open in design view ---
        try:
            _run_with_timeout(lambda name=form_name: app.DoCmd.OpenForm(name, AC_VIEW_DESIGN))
        except COMTimeoutError as exc:
            warnings.append(f"Timed out opening form '{form_name}': {exc}")
            _cleanup_file(out_file)
            return forms  # Access was killed; abort remaining forms
        except Exception as exc:
            warnings.append(f"Could not open form '{form_name}' in design view: {exc}")
            _cleanup_file(out_file)
            continue

        try:
            frm = app.Forms(form_name)
            record_source = _safe_getattr(frm, "RecordSource")

            controls: list[dict] = []
            try:
                for j in range(frm.Controls.Count):
                    ctrl = frm.Controls.Item(j)
                    controls.append(_extract_control(ctrl))
            except Exception as exc:
                warnings.append(f"Error reading controls of form '{form_name}': {exc}")

            forms.append({
                "name": form_name,
                "record_source": record_source,
                "controls": controls,
            })

        except Exception as exc:
            warnings.append(f"Error extracting form '{form_name}': {exc}")
        finally:
            try:
                app.DoCmd.Close(AC_FORM, form_name, AC_SAVE_NO)
            except Exception:
                pass
            _cleanup_file(out_file)

    shutil.rmtree(tmp_dir, ignore_errors=True)
    return forms


# ---------------------------------------------------------------------------
# Report extraction
# ---------------------------------------------------------------------------


def extract_reports(app, warnings: list[str]) -> list[dict]:
    """Extract all reports using SaveAsText (primary) or design view (fallback)."""
    reports: list[dict] = []

    try:
        all_reports = app.CurrentProject.AllReports
    except Exception as exc:
        warnings.append(f"Could not enumerate reports: {exc}")
        return reports

    report_count = all_reports.Count
    print(f"  Found {report_count} report(s)")

    tmp_dir = Path(tempfile.mkdtemp(prefix="access_reports_"))

    for i in range(report_count):
        report_obj = all_reports.Item(i)
        report_name = report_obj.Name
        print(f"    Processing report: {report_name}")

        # --- Primary method: SaveAsText ---
        out_file = tmp_dir / f"{report_name}.txt"
        try:
            app.SaveAsText(AC_REPORT, report_name, str(out_file))
            report_text = _read_access_text_file(out_file)
            parsed = _parse_form_text(report_text, report_name, is_report=True)
            if parsed:
                reports.append(parsed)
                _cleanup_file(out_file)
                continue
        except Exception:
            pass

        # --- Fallback: open in design view ---
        try:
            _run_with_timeout(lambda name=report_name: app.DoCmd.OpenReport(name, AC_VIEW_DESIGN))
        except COMTimeoutError as exc:
            warnings.append(f"Timed out opening report '{report_name}': {exc}")
            _cleanup_file(out_file)
            return reports  # Access was killed; abort remaining reports
        except Exception as exc:
            warnings.append(f"Could not open report '{report_name}' in design view: {exc}")
            _cleanup_file(out_file)
            continue

        try:
            rpt = app.Reports(report_name)
            record_source = _safe_getattr(rpt, "RecordSource")

            controls: list[dict] = []
            try:
                for j in range(rpt.Controls.Count):
                    ctrl = rpt.Controls.Item(j)
                    controls.append(_extract_control(ctrl))
            except Exception as exc:
                warnings.append(f"Error reading controls of report '{report_name}': {exc}")

            group_levels: list[dict] = []
            try:
                gl_count = rpt.GroupLevel.Count if hasattr(rpt, "GroupLevel") else 0
                for g in range(gl_count):
                    gl = rpt.GroupLevel(g)
                    group_levels.append({
                        "field": _safe_getattr(gl, "ControlSource", ""),
                        "group_on": _safe_getattr(gl, "GroupOn", 0),
                        "sort_order": _safe_getattr(gl, "SortOrder", 0),
                    })
            except Exception:
                pass

            reports.append({
                "name": report_name,
                "record_source": record_source,
                "controls": controls,
                "group_levels": group_levels,
            })

        except Exception as exc:
            warnings.append(f"Error extracting report '{report_name}': {exc}")
        finally:
            try:
                app.DoCmd.Close(AC_REPORT, report_name, AC_SAVE_NO)
            except Exception:
                pass
            _cleanup_file(out_file)

    shutil.rmtree(tmp_dir, ignore_errors=True)
    return reports


# ---------------------------------------------------------------------------
# VBA module extraction
# ---------------------------------------------------------------------------


def extract_vba_modules(app, warnings: list[str]) -> list[dict]:
    """Extract VBA source code from all modules in the VBA project."""
    modules: list[dict] = []

    try:
        vb_project = app.VBE.ActiveVBProject
    except Exception as exc:
        error_str = str(exc)
        if "6068" in error_str or "programmatic access" in error_str.lower():
            warnings.append(
                "VBE access denied (error 6068). To enable: "
                "File → Options → Trust Center → Trust Center Settings → Macro Settings → "
                "check 'Trust access to the VBA project object model'. "
                "VBA source code will NOT be extracted."
            )
        else:
            warnings.append(f"Could not access VBA project: {exc}")
        return modules

    print("  Extracting VBA modules...")
    try:
        components = vb_project.VBComponents
        for i in range(1, components.Count + 1):  # VBComponents is 1-indexed
            comp = components.Item(i)
            comp_name = comp.Name
            comp_type = comp.Type
            type_name = VB_COMPONENT_TYPES.get(comp_type, f"Unknown({comp_type})")

            line_count = 0
            source = ""
            try:
                code_module = comp.CodeModule
                line_count = code_module.CountOfLines
                if line_count > 0:
                    source = code_module.Lines(1, line_count)
            except Exception as exc:
                warnings.append(f"Could not read source of module '{comp_name}': {exc}")

            if line_count > 5000:
                warnings.append(
                    f"Large VBA module: '{comp_name}' has {line_count} lines"
                )

            has_api = bool(RE_API_DECLARATION.search(source)) if source else False
            external_refs = sorted(set(RE_EXTERNAL_REF.findall(source))) if source else []

            modules.append({
                "name": comp_name,
                "type": type_name,
                "type_id": comp_type,
                "line_count": line_count,
                "source": source,
                "has_api_declarations": has_api,
                "external_references": external_refs,
            })
            print(f"    Module: {comp_name} ({type_name}, {line_count} lines)")

    except Exception as exc:
        warnings.append(f"Error iterating VBA components: {exc}")

    return modules


# ---------------------------------------------------------------------------
# Macro extraction
# ---------------------------------------------------------------------------


def extract_macros(app, warnings: list[str]) -> list[dict]:
    """Extract macros using SaveAsText to get the full macro definition."""
    macros: list[dict] = []

    try:
        all_macros = app.CurrentProject.AllMacros
    except Exception as exc:
        warnings.append(f"Could not enumerate macros: {exc}")
        return macros

    macro_count = all_macros.Count
    print(f"  Found {macro_count} macro(s)")

    tmp_dir = Path(tempfile.mkdtemp(prefix="access_macros_"))

    for i in range(macro_count):
        macro_obj = all_macros.Item(i)
        macro_name = macro_obj.Name
        print(f"    Extracting macro: {macro_name}")

        raw_text = ""
        out_file = tmp_dir / f"{macro_name}.txt"
        try:
            app.SaveAsText(AC_MACRO, macro_name, str(out_file))
            raw_text = _read_access_text_file(out_file)
        except Exception as exc:
            warnings.append(f"SaveAsText failed for macro '{macro_name}': {exc}")

        _cleanup_file(out_file)

        macros.append({
            "name": macro_name,
            "raw_text": raw_text,
        })

    shutil.rmtree(tmp_dir, ignore_errors=True)

    return macros


# ---------------------------------------------------------------------------
# Classification metrics
# ---------------------------------------------------------------------------


def compute_metrics(
    forms: list[dict],
    reports: list[dict],
    vba_modules: list[dict],
    macros: list[dict],
) -> dict:
    """Compute complexity classification metrics from extracted data."""
    vba_total_lines = sum(m["line_count"] for m in vba_modules)
    vba_module_count = len(vba_modules)
    has_api = any(m["has_api_declarations"] for m in vba_modules)
    has_external = any(len(m["external_references"]) > 0 for m in vba_modules)

    form_count = len(forms)
    report_count = len(reports)
    macro_count = len(macros)

    subform_count = 0
    event_procedure_count = 0
    combo_box_count = 0

    for obj in [*forms, *reports]:
        for ctrl in obj.get("controls", []):
            if ctrl.get("type") == "Subform":
                subform_count += 1
            if ctrl.get("type") in ("ComboBox", "ListBox"):
                combo_box_count += 1
            event_procedure_count += len(ctrl.get("events", []))

    # Weighted complexity score
    score = 0

    # VBA lines: 0-50=0, 50-500=1, 500-2000=3, 2000+=5
    if vba_total_lines > 2000:
        score += 5
    elif vba_total_lines > 500:
        score += 3
    elif vba_total_lines > 50:
        score += 1

    # Forms: 0-2=0, 3-5=1, 6+=2
    if form_count >= 6:
        score += 2
    elif form_count >= 3:
        score += 1

    # Subforms: each adds 1pt
    score += subform_count

    # API declarations: +3
    if has_api:
        score += 3

    # External objects: +2
    if has_external:
        score += 2

    # Event procedures: 0-5=0, 5-20=1, 20+=3
    if event_procedure_count > 20:
        score += 3
    elif event_procedure_count > 5:
        score += 1

    # Determine tier
    if score >= 7:
        tier = "complex"
    elif score >= 3:
        tier = "moderate"
    else:
        tier = "simple"

    return {
        "vba_total_lines": vba_total_lines,
        "vba_module_count": vba_module_count,
        "has_api_declarations": has_api,
        "has_external_objects": has_external,
        "form_count": form_count,
        "report_count": report_count,
        "macro_count": macro_count,
        "subform_count": subform_count,
        "event_procedure_count": event_procedure_count,
        "combo_box_count": combo_box_count,
        "complexity_score": score,
        "complexity_tier": tier,
    }


# ---------------------------------------------------------------------------
# Output generation
# ---------------------------------------------------------------------------


def write_json(output_dir: Path, forms, reports, vba_modules, macros, metrics) -> Path:
    """Write the structured JSON output."""
    data = {
        "forms": forms,
        "reports": reports,
        "vba_modules": vba_modules,
        "macros": macros,
        "metrics": metrics,
    }
    out_path = output_dir / "forms_vba.json"
    out_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return out_path


def write_summary_md(
    output_dir: Path,
    forms: list[dict],
    reports: list[dict],
    vba_modules: list[dict],
    macros: list[dict],
    metrics: dict,
    warnings: list[str],
) -> Path:
    """Write a human-readable markdown summary."""
    lines: list[str] = []
    lines.append("# Forms & VBA Extraction Summary\n")

    # Complexity overview
    lines.append("## Complexity\n")
    lines.append(f"- **Tier**: {metrics['complexity_tier']}")
    lines.append(f"- **Score**: {metrics['complexity_score']}")
    lines.append(f"- **VBA lines**: {metrics['vba_total_lines']}")
    lines.append(f"- **Event procedures**: {metrics['event_procedure_count']}")
    lines.append("")

    # Forms
    lines.append(f"## Forms ({metrics['form_count']})\n")
    if forms:
        lines.append("| Form | Record Source | Controls | Events |")
        lines.append("|------|--------------|----------|--------|")
        for f in forms:
            ctrl_count = len(f.get("controls", []))
            event_count = sum(len(c.get("events", [])) for c in f.get("controls", []))
            rs = (f.get("record_source") or "—").replace("|", "\\|")
            lines.append(f"| {f['name']} | {rs} | {ctrl_count} | {event_count} |")
        lines.append("")
    else:
        lines.append("No forms found.\n")

    # Reports
    lines.append(f"## Reports ({metrics['report_count']})\n")
    if reports:
        lines.append("| Report | Record Source | Controls |")
        lines.append("|--------|--------------|----------|")
        for r in reports:
            ctrl_count = len(r.get("controls", []))
            rs = (r.get("record_source") or "—").replace("|", "\\|")
            lines.append(f"| {r['name']} | {rs} | {ctrl_count} |")
        lines.append("")
    else:
        lines.append("No reports found.\n")

    # VBA Modules
    lines.append(f"## VBA Modules ({metrics['vba_module_count']})\n")
    if vba_modules:
        lines.append("| Module | Type | Lines | API Decl | External Refs |")
        lines.append("|--------|------|-------|----------|---------------|")
        for m in vba_modules:
            ext = ", ".join(m["external_references"]) if m["external_references"] else "—"
            api = "Yes" if m["has_api_declarations"] else "No"
            lines.append(f"| {m['name']} | {m['type']} | {m['line_count']} | {api} | {ext} |")
        lines.append("")
    else:
        lines.append("No VBA modules found.\n")

    # Macros
    lines.append(f"## Macros ({metrics['macro_count']})\n")
    if macros:
        for m in macros:
            lines.append(f"- {m['name']}")
        lines.append("")
    else:
        lines.append("No macros found.\n")

    # Warnings
    if warnings:
        lines.append("## Warnings\n")
        for w in warnings:
            lines.append(f"- ⚠️ {w}")
        lines.append("")

    out_path = output_dir / "forms_vba_summary.md"
    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path


# ---------------------------------------------------------------------------
# Main extraction orchestrator
# ---------------------------------------------------------------------------


def extract_database(db_path: Path, output_dir: Path, password: str | None = None) -> None:
    """Run full forms/VBA extraction for a single database."""
    if not db_path.exists():
        raise FileNotFoundError(f"Access file not found: {db_path}")

    print(f"\n{'='*60}")
    print(f"Extracting forms & VBA: {db_path.name}")
    print(f"{'='*60}")

    output_dir.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []

    app = open_access_app(db_path, password)
    forms: list[dict] = []
    reports: list[dict] = []
    vba_modules: list[dict] = []
    macros: list[dict] = []

    try:
        print("\n  [1/4] Extracting forms...")
        forms = extract_forms(app, warnings)

        print("\n  [2/4] Extracting reports...")
        reports = extract_reports(app, warnings)

        print("\n  [3/4] Extracting VBA modules...")
        vba_modules = extract_vba_modules(app, warnings)

        print("\n  [4/4] Extracting macros...")
        macros = extract_macros(app, warnings)

    except Exception as exc:
        warnings.append(f"Extraction aborted early: {exc}")
    finally:
        close_access_app(app)

    # Compute metrics
    metrics = compute_metrics(forms, reports, vba_modules, macros)

    # Write outputs
    json_path = write_json(output_dir, forms, reports, vba_modules, macros, metrics)
    md_path = write_summary_md(output_dir, forms, reports, vba_modules, macros, metrics, warnings)

    # Print summary
    print(f"\n  Results:")
    print(f"    Forms:      {metrics['form_count']}")
    print(f"    Reports:    {metrics['report_count']}")
    print(f"    VBA lines:  {metrics['vba_total_lines']}")
    print(f"    Macros:     {metrics['macro_count']}")
    print(f"    Complexity: {metrics['complexity_tier']} (score={metrics['complexity_score']})")
    if warnings:
        print(f"    Warnings:   {len(warnings)}")
    print(f"\n  Output: {json_path}")
    print(f"          {md_path}")


def main() -> None:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    access_dir = Path(args.access_dir).resolve()
    access_dir.mkdir(parents=True, exist_ok=True)

    db_paths = resolve_db_paths(args, access_dir)
    multi_db = len(db_paths) > 1

    for db_path in db_paths:
        output_dir = resolve_output_dir(repo_root, args.output_dir, db_path, multi_db)
        extract_database(db_path, output_dir, args.password)

    print("\nDone.")


if __name__ == "__main__":
    main()
