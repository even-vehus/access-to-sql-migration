"""
Artifact Inspector
Reads schema.json + forms_vba.json from a migration_output folder and produces:
  - app_spec.json  — compact machine-readable specification for code generation
  - app_spec.md    — human-readable overview with VBA classification and form→entity mapping

Classification is driven by *generic structural/content heuristics* (FK topology,
column counts, VBA body content + naming conventions) so it works on any Access
database, not just the bundled Northwind sample. Genuine product decisions that
no heuristic can infer (which entities are "v1 core", the app title, preferred
C# class names, custom routes) can be pinned in an optional override config:

  - pass  --config path/to/overrides.json, or
  - drop  generators/configs/<db-name>.json  (auto-loaded)

Usage:
    python generators/inspect_artifacts.py
    python generators/inspect_artifacts.py --db-name NorthwindStarterED
    python generators/inspect_artifacts.py --output-dir path/to/custom_dir
    python generators/inspect_artifacts.py --db-name MyApp --config overrides.json
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

# Tables with at most this many columns that neither reference nor are referenced
# by anything are treated as utility/config islands ("later"), not core entities.
ISOLATED_MAX_COLS = 12


# ---------------------------------------------------------------------------
# Small string helpers
# ---------------------------------------------------------------------------


def _to_pascal(name: str) -> str:
    parts = re.split(r"[_\s]+", name)
    return "".join(p[:1].upper() + p[1:] for p in parts if p)


def _kebab(name: str) -> str:
    """CamelCase / snake_case → kebab-case (e.g. PurchaseOrders → purchase-orders)."""
    s = re.sub(r"(?<!^)(?=[A-Z])", "-", name)
    s = re.sub(r"[_\s]+", "-", s)
    return re.sub(r"-+", "-", s).lower().strip("-")


def _norm_entity_key(word: str) -> str:
    """Normalise a name for singular/plural-tolerant matching.

    Applied to BOTH sides of a comparison, so mild over-stripping is harmless as
    long as it is consistent (e.g. 'Companies'→'company', 'Company'→'company')."""
    w = word.strip().lower()
    if w.endswith("ies"):
        w = w[:-3] + "y"
    elif w.endswith("ses"):
        w = w[:-2]
    elif w.endswith("s") and not w.endswith("ss"):
        w = w[:-1]
    return w


# ---------------------------------------------------------------------------
# Type mapping: Access/T-SQL → C# type
# ---------------------------------------------------------------------------

_TSQL_TO_CSHARP: dict[str, str] = {
    "LONG": "int",
    "INTEGER": "int",
    "INT": "int",
    "SMALLINT": "short",
    "BYTE": "byte",
    "DOUBLE": "double",
    "SINGLE": "float",
    "DECIMAL": "decimal",
    "CURRENCY": "decimal",
    "NVARCHAR": "string",
    "NVARCHAR(MAX)": "string",
    "VARCHAR": "string",
    "CHAR": "string",
    "TEXT": "string",
    "DATETIME": "DateTime",
    "DATETIME2": "DateTime",
    "DATE": "DateOnly",
    "BOOLEAN": "bool",
    "BIT": "bool",
    "BINARY": "byte[]",
    "VARBINARY": "byte[]",
    "UNIQUEIDENTIFIER": "Guid",
}


def _to_csharp_type(access_type: str, nullable: bool) -> str:
    base = _TSQL_TO_CSHARP.get(access_type.upper(), "object")
    if nullable and base not in ("string", "byte[]", "object"):
        return f"{base}?"
    return base


# ---------------------------------------------------------------------------
# Override config
# ---------------------------------------------------------------------------


def load_config(db_name: str, repo_root: Path, explicit_path: str | None) -> dict:
    """Load an optional override config (JSON). Precedence:
    1. --config <path>
    2. generators/configs/<db_name>.json
    Returns {} when none is found."""
    candidates: list[Path] = []
    if explicit_path:
        candidates.append(Path(explicit_path))
    candidates.append(repo_root / "generators" / "configs" / f"{db_name}.json")

    for path in candidates:
        if path and path.exists():
            try:
                cfg = json.loads(path.read_text(encoding="utf-8"))
                print(f"  Using override config: {path}")
                return cfg
            except Exception as exc:  # noqa: BLE001 — config is best-effort
                print(f"  [config warning] could not read {path}: {exc}")
    return {}


# ---------------------------------------------------------------------------
# VBA module classification (content + naming convention, not a name allow-list)
# ---------------------------------------------------------------------------

# A module whose *whole* name (sans mod/cls/bas prefix) is one of these is generic
# plumbing replaced by the framework (EF Core, ILogger, System.IO, …).
_INFRA_NAME = re.compile(
    r"(mod|cls|bas)?(dao|ado|dataaccess|db|database|file|files|fileio|io|log|logger"
    r"|logging|debug|trace|registry|interop|win32|winapi|api)",
    re.IGNORECASE,
)
# A module whose name clearly marks it as UI orchestration.
_UI_NAME = re.compile(
    r"(mod|cls|bas)?(forms?|ribbon|ribboncallback|startup|global|navigation|nav|menu|ui)",
    re.IGNORECASE,
)
# Win32 API declaration inside the body → infrastructure.
_DECLARE = re.compile(
    r"(?im)^\s*(public|private)?\s*declare\s+(ptrsafe\s+)?(function|sub)\b"
)
# Access UI automation inside the body → ui_glue.
_UI_BODY = re.compile(r"\bDoCmd\b|\bForms?\s*!|\bReports?\s*!|\bScreen\s*\.")


def _classify_module(name: str, module_type: str, source: str = "", has_api: bool = False) -> str:
    if module_type == "Document":
        return "ui_glue"  # Form_* / Report_* code-behind

    src = source or ""

    # Infrastructure. Content alone is a poor signal here (domain modules use DAO
    # too), so beyond Win32 API declarations we rely on a naming convention.
    if has_api or _DECLARE.search(src):
        return "infrastructure"
    if _INFRA_NAME.fullmatch(name):
        return "infrastructure"

    # UI glue: Access UI automation in the body, or an unmistakably UI module name.
    if _UI_BODY.search(src) or _UI_NAME.fullmatch(name):
        return "ui_glue"

    # Default: portable business logic (incl. non-UI class modules).
    return "domain"


def _suggest_csharp_class(name: str, classification: str) -> str | None:
    if classification != "domain":
        return None
    base = re.sub(r"(?i)^(mod|cls|bas)", "", name) or name
    pas = _to_pascal(base)
    if re.search(r"(?i)(helper|service|manager|handler|util|utils|utilities)$", pas):
        return pas
    return f"{pas}Service"


# ---------------------------------------------------------------------------
# Entity scope classification (structural)
# ---------------------------------------------------------------------------


def _classify_entity(name: str, n_cols: int, out_fk: int, in_deg: int) -> str:
    """Default, structural scope. Overridable per-table via config['entity_scope'].

    - Access system tables (USys*/MSys*)            → later
    - isolated small tables (no FK in or out)        → later (utility/config islands)
    - leaf tables referenced by others (out_fk == 0) → v1_support (lookup/dimension)
    - FK-heavy tables nothing references (junction)  → v1_support
    - everything else (real, connected entities)     → v1_core
    """
    if re.match(r"(?i)^(u|m)sys", name):
        return "later"
    if out_fk == 0 and in_deg == 0:
        return "later" if n_cols <= ISOLATED_MAX_COLS else "v1_core"
    if out_fk == 0 and in_deg >= 1:
        return "v1_support"
    if out_fk >= 2 and in_deg == 0 and n_cols <= out_fk + 2:
        return "v1_support"
    return "v1_core"


# ---------------------------------------------------------------------------
# Schema analysis
# ---------------------------------------------------------------------------


def analyse_schema(schema: dict, config: dict | None = None) -> dict:
    config = config or {}
    scope_overrides: dict = config.get("entity_scope", {})

    tables = {k: v for k, v in schema.items() if isinstance(v, dict) and not k.startswith("_")}

    # Pre-compute FK in-degree (how many *other* tables reference each table).
    in_degree: dict[str, int] = {t: 0 for t in tables}
    for table_name, table_data in tables.items():
        for fk in table_data.get("foreign_keys", []):
            ref = fk.get("pk_table")
            if ref in in_degree and ref != table_name:
                in_degree[ref] += 1

    entities = {}
    fk_graph: dict[str, list[dict]] = {}

    for table_name, table_data in tables.items():
        columns = table_data.get("columns", [])
        pks = set(table_data.get("primary_keys", []))
        fks = table_data.get("foreign_keys", [])
        row_count = table_data.get("row_count", 0)

        entity_name = _to_pascal(table_name)
        out_fk = len(fks)
        scope = scope_overrides.get(
            table_name, _classify_entity(table_name, len(columns), out_fk, in_degree[table_name])
        )

        fields = []
        for col in columns:
            col_name = col["name"]
            csharp_type = _to_csharp_type(col["type_name"], col.get("nullable", True))
            fields.append({
                "name": col_name,
                "csharp_name": _to_pascal(col_name),
                "access_type": col["type_name"],
                "csharp_type": csharp_type,
                "is_pk": col_name in pks,
                "is_fk": any(fk["fk_column"] == col_name for fk in fks),
                "nullable": col.get("nullable", True),
                "max_length": col.get("column_size"),
            })

        fk_graph[table_name] = [
            {
                "from_column": fk["fk_column"],
                "to_table": fk["pk_table"],
                "to_column": fk["pk_column"],
            }
            for fk in fks
        ]

        entities[table_name] = {
            "table_name": table_name,
            "entity_name": entity_name,
            "scope": scope,
            "row_count": row_count,
            "fields": fields,
            "foreign_keys": fk_graph[table_name],
            "primary_keys": list(pks),
        }

    return {"entities": entities, "fk_graph": fk_graph}


# ---------------------------------------------------------------------------
# Form/Report → entity matching
# ---------------------------------------------------------------------------

# Leading object-type prefixes and trailing role words stripped when deriving an
# entity name from a form/report name.
_NAME_PREFIX = re.compile(r"(?i)^(s?frm|s?rpt|dlg|frm|rpt)")
_ROLE_SUFFIX = re.compile(
    r"(?i)(List|Details?|Edit|Editor|New|Form|Manager|Mgr|View|Card|Dialog|Subform|Sub)$"
)


def _resolve_entity(token: str | None, entities: dict) -> str | None:
    if not token or not token.strip():
        return None
    tl = token.strip().lower()
    for name in entities:  # exact (case-insensitive)
        if name.lower() == tl:
            return name
    key = _norm_entity_key(token)  # singular/plural tolerant
    for name in entities:
        if _norm_entity_key(name) == key:
            return name
    return None


def _match_form_to_entity(name: str, entities: dict) -> str | None:
    """Best-effort entity match from a form/report name (used when there is no
    record source). e.g. frmCompanyList→Companies, frmOrderDetails→Orders."""
    base = _NAME_PREFIX.sub("", name)
    candidates = [_ROLE_SUFFIX.sub("", base)]
    for part in re.split(r"[_\s]+", base):
        candidates.append(_ROLE_SUFFIX.sub("", part))
        candidates.append(part)
    for cand in candidates:
        entity = _resolve_entity(cand, entities)
        if entity:
            return entity
    return None


def _form_to_entity(record_source: str | None, name: str, entities: dict) -> str | None:
    # 1. Record source: direct table, or "SELECT ... FROM <table>".
    if record_source:
        rs = record_source.strip().strip('"')
        if rs in entities:
            return rs
        m = re.search(r"\bFROM\s+\[?(\w+)\]?", rs, re.IGNORECASE)
        if m and m.group(1) in entities:
            return m.group(1)
        entity = _resolve_entity(rs, entities)
        if entity:
            return entity
    # 2. Fall back to the form/report name (Access forms often set RecordSource at runtime).
    return _match_form_to_entity(name, entities)


def _is_dialog(name: str) -> bool:
    return bool(re.search(r"(?i)(dialog|login|credential|confirm|prompt|picker|msgbox)", name))


def _suggest_route(form_name: str, entity: str | None, is_subform: bool, is_dialog: bool) -> str | None:
    if is_subform or is_dialog:
        return None  # subforms render as child components; dialogs as modals
    if entity:
        slug = _kebab(entity)
        if re.search(r"(?i)detail", form_name):
            return f"/{slug}/:id"
        return f"/{slug}"
    # No entity: only forms that look like pages get a route.
    if re.search(r"(?i)(list|detail|manager|board|home|dashboard|search)", form_name):
        return f"/{_kebab(_NAME_PREFIX.sub('', form_name))}"
    return None


def _suggest_report_component(report_name: str) -> str:
    base = _NAME_PREFIX.sub("", report_name)
    pas = _to_pascal(base) or _to_pascal(report_name)
    return pas if pas.lower().endswith("report") else f"{pas}Report"


def analyse_forms(forms_vba: dict, entities: dict, config: dict | None = None) -> dict:
    config = config or {}
    route_overrides: dict = config.get("routes", {})
    component_overrides: dict = config.get("report_components", {})

    forms_info = []
    for form in forms_vba.get("forms", []):
        name = form["name"]
        rs = form.get("record_source")
        entity = _form_to_entity(rs, name, entities)
        control_count = len(form.get("controls", []))
        event_count = sum(len(c.get("events", [])) for c in form.get("controls", []))
        is_subform = name.startswith("sfrm")
        is_dialog = _is_dialog(name)
        route = (
            route_overrides[name]
            if name in route_overrides
            else _suggest_route(name, entity, is_subform, is_dialog)
        )
        forms_info.append({
            "name": name,
            "record_source": rs,
            "mapped_entity": entity,
            "control_count": control_count,
            "event_count": event_count,
            "is_subform": is_subform,
            "is_dialog": is_dialog,
            "suggested_route": route,
        })

    reports_info = []
    for report in forms_vba.get("reports", []):
        name = report["name"]
        rs = report.get("record_source")
        entity = _form_to_entity(rs, name, entities)
        component = component_overrides.get(name) or _suggest_report_component(name)
        reports_info.append({
            "name": name,
            "record_source": rs,
            "mapped_entity": entity,
            "control_count": len(report.get("controls", [])),
            "suggested_component": component,
        })

    return {"forms": forms_info, "reports": reports_info}


# ---------------------------------------------------------------------------
# VBA module analysis
# ---------------------------------------------------------------------------


def analyse_vba_modules(forms_vba: dict, config: dict | None = None) -> list[dict]:
    config = config or {}
    class_overrides: dict = config.get("vba_classification", {})
    csharp_overrides: dict = config.get("csharp_class", {})

    modules = []
    for mod in forms_vba.get("vba_modules", []):
        name = mod["name"]
        mod_type = mod.get("type", "")
        classification = class_overrides.get(name) or _classify_module(
            name, mod_type, mod.get("source", ""), mod.get("has_api_declarations", False)
        )
        csharp_class = csharp_overrides.get(name) or _suggest_csharp_class(name, classification)
        modules.append({
            "name": name,
            "type": mod_type,
            "line_count": mod.get("line_count", 0),
            "classification": classification,
            "has_api_declarations": mod.get("has_api_declarations", False),
            "external_references": mod.get("external_references", []),
            "suggested_csharp_class": csharp_class,
        })
    # Sort: domain first, then ui_glue, infra, unknown
    order = {"domain": 0, "ui_glue": 1, "infrastructure": 2, "unknown": 3}
    modules.sort(key=lambda m: (order.get(m["classification"], 3), m["name"]))
    return modules


# ---------------------------------------------------------------------------
# Markdown generation
# ---------------------------------------------------------------------------


def build_markdown(
    title: str,
    schema_analysis: dict,
    form_analysis: dict,
    vba_modules: list[dict],
    metrics: dict,
) -> str:
    lines: list[str] = []
    lines.append(f"# App Specification — {title}\n")
    lines.append("Generated by `generators/inspect_artifacts.py`. Use as input for code generation.\n")

    # --- Complexity ---
    lines.append("## Complexity\n")
    lines.append(f"- **Tier**: {metrics.get('complexity_tier', '—')}")
    lines.append(f"- **Score**: {metrics.get('complexity_score', '—')}")
    lines.append(f"- **VBA lines**: {metrics.get('vba_total_lines', '—')}")
    lines.append(f"- **Forms**: {metrics.get('form_count', '—')}")
    lines.append(f"- **Reports**: {metrics.get('report_count', '—')}")
    lines.append("")

    # --- Entities ---
    lines.append("## Entities (tables → C# classes)\n")

    for scope_label, scope_key in [("v1 Core", "v1_core"), ("v1 Support (lookup tables)", "v1_support"), ("Later", "later")]:
        scoped = [e for e in schema_analysis["entities"].values() if e["scope"] == scope_key]
        if not scoped:
            continue
        lines.append(f"### {scope_label}\n")
        lines.append("| Table | Entity | Rows | PK | Foreign Keys |")
        lines.append("|-------|--------|------|----|--------------|")
        for e in scoped:
            pk_str = ", ".join(e["primary_keys"])
            fk_str = "; ".join(f"{fk['from_column']} → {fk['to_table']}.{fk['to_column']}" for fk in e["foreign_keys"])
            lines.append(f"| {e['table_name']} | {e['entity_name']} | {e['row_count']} | {pk_str} | {fk_str or '—'} |")
        lines.append("")

    # --- Entity field details for v1 core ---
    lines.append("## Entity Field Details (v1 Core only)\n")
    for e in schema_analysis["entities"].values():
        if e["scope"] != "v1_core":
            continue
        lines.append(f"### {e['entity_name']}\n")
        lines.append("| Column | C# Name | C# Type | PK | FK | Nullable |")
        lines.append("|--------|---------|---------|----|----|----------|")
        for f in e["fields"]:
            pk = "✓" if f["is_pk"] else ""
            fk = "✓" if f["is_fk"] else ""
            nullable = "✓" if f["nullable"] else ""
            lines.append(f"| {f['name']} | {f['csharp_name']} | `{f['csharp_type']}` | {pk} | {fk} | {nullable} |")
        lines.append("")

    # --- Forms → Routes ---
    lines.append("## Forms → React Routes\n")
    routed_forms = [f for f in form_analysis["forms"] if f.get("suggested_route")]
    subforms = [f for f in form_analysis["forms"] if f["is_subform"]]
    dialogs = [f for f in form_analysis["forms"] if f["is_dialog"] and not f["is_subform"]]

    if routed_forms:
        lines.append("### Main Forms (each becomes a React route)\n")
        lines.append("| Access Form | Route | Entity | Controls |")
        lines.append("|-------------|-------|--------|----------|")
        for f in routed_forms:
            lines.append(f"| {f['name']} | `{f['suggested_route']}` | {f['mapped_entity'] or '—'} | {f['control_count']} |")
        lines.append("")

    if subforms:
        lines.append("### Subforms (become React child components)\n")
        lines.append("| Subform | Entity | Controls |")
        lines.append("|---------|--------|----------|")
        for f in subforms:
            lines.append(f"| {f['name']} | {f['mapped_entity'] or '—'} | {f['control_count']} |")
        lines.append("")

    if dialogs:
        lines.append("### Dialogs (become MUI Dialog components)\n")
        for f in dialogs:
            lines.append(f"- {f['name']}")
        lines.append("")

    # --- Reports ---
    lines.append("## Reports → React Report Pages\n")
    lines.append("| Access Report | React Component | Entity |")
    lines.append("|---------------|-----------------|--------|")
    for r in form_analysis["reports"]:
        if r["name"].startswith("_"):
            continue
        lines.append(f"| {r['name']} | `{r['suggested_component']}` | {r['mapped_entity'] or '—'} |")
    lines.append("")

    # --- VBA modules ---
    lines.append("## VBA Module Classification\n")

    for cls_label, cls_key in [
        ("Domain → Port to C#", "domain"),
        ("UI Glue → React state/hooks (skip VBA)", "ui_glue"),
        ("Infrastructure → Replaced by framework", "infrastructure"),
        ("Unknown", "unknown"),
    ]:
        mods = [m for m in vba_modules if m["classification"] == cls_key]
        if not mods:
            continue
        lines.append(f"### {cls_label}\n")
        lines.append("| Module | Type | Lines | C# Class |")
        lines.append("|--------|------|-------|----------|")
        for m in mods:
            csharp = m.get("suggested_csharp_class") or "—"
            lines.append(f"| {m['name']} | {m['type']} | {m['line_count']} | {csharp} |")
        lines.append("")

    # --- VBA port queue (smallest/foundational domain modules first) ---
    domain_mods = [m for m in vba_modules if m["classification"] == "domain"]
    if domain_mods:
        lines.append("## VBA Port Queue (recommended order)\n")
        lines.append("Smallest domain modules first — port foundational helpers before larger logic.\n")
        for i, mod in enumerate(sorted(domain_mods, key=lambda m: (m["line_count"], m["name"])), 1):
            lines.append(f"{i}. `{mod['name']}` ({mod['line_count']} lines) → `{mod['suggested_csharp_class']}`")
        lines.append("")

    # --- FK dependency order for migrations ---
    lines.append("## Table Insertion Order (FK-safe)\n")
    lines.append("Useful for seeding and deployment. Leaf tables first.\n")
    order = _fk_topological_sort(schema_analysis["entities"])
    for idx, tname in enumerate(order, 1):
        lines.append(f"{idx}. {tname}")
    lines.append("")

    return "\n".join(lines)


def _fk_topological_sort(entities: dict) -> list[str]:
    """Kahn's algorithm — returns tables in FK-dependency order (dependencies first)."""
    in_degree = {t: 0 for t in entities}
    adj: dict[str, set[str]] = {t: set() for t in entities}

    for table_name, entity in entities.items():
        for fk in entity.get("foreign_keys", []):
            ref = fk.get("to_table") or fk.get("pk_table", "")
            if ref in entities and ref != table_name:
                if table_name not in adj[ref]:
                    adj[ref].add(table_name)
                    in_degree[table_name] += 1

    queue = [t for t, d in in_degree.items() if d == 0]
    queue.sort()
    result = []
    while queue:
        node = queue.pop(0)
        result.append(node)
        for neighbour in sorted(adj[node]):
            in_degree[neighbour] -= 1
            if in_degree[neighbour] == 0:
                queue.append(neighbour)

    # Append any remaining (cycles)
    remaining = [t for t in entities if t not in result]
    result.extend(sorted(remaining))
    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Produce app_spec.json + app_spec.md from schema.json and forms_vba.json."
    )
    parser.add_argument(
        "--db-name",
        default=None,
        help=(
            "Sub-folder name under migration_output (default: auto-detect or process all). "
            "Example: NorthwindStarterED"
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Override output directory (default: migration_output/<db-name>).",
    )
    parser.add_argument(
        "--config",
        default=None,
        help=(
            "Optional override config (JSON) pinning entity_scope, title, routes, "
            "report_components, vba_classification, csharp_class. "
            "Auto-loaded from generators/configs/<db-name>.json when present."
        ),
    )
    return parser.parse_args()


def find_db_dirs(output_root: Path) -> list[Path]:
    return sorted(
        p for p in output_root.iterdir()
        if p.is_dir() and (p / "schema.json").exists() and (p / "forms_vba.json").exists()
    )


def process_db(db_dir: Path, repo_root: Path, explicit_config: str | None) -> None:
    print(f"\n{'='*60}")
    print(f"Inspecting: {db_dir.name}")
    print(f"{'='*60}")

    schema_path = db_dir / "schema.json"
    forms_vba_path = db_dir / "forms_vba.json"

    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    forms_vba = json.loads(forms_vba_path.read_text(encoding="utf-8"))

    config = load_config(db_dir.name, repo_root, explicit_config)
    title = config.get("title") or f"{db_dir.name} Modernization"

    schema_analysis = analyse_schema(schema, config)
    form_analysis = analyse_forms(forms_vba, schema_analysis["entities"], config)
    vba_modules = analyse_vba_modules(forms_vba, config)
    metrics = forms_vba.get("metrics", {})

    app_spec = {
        "db_name": db_dir.name,
        "title": title,
        "schema": schema_analysis,
        "forms": form_analysis["forms"],
        "reports": form_analysis["reports"],
        "vba_modules": vba_modules,
        "metrics": metrics,
    }

    spec_json_path = db_dir / "app_spec.json"
    spec_json_path.write_text(json.dumps(app_spec, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  Written: {spec_json_path}")

    md = build_markdown(title, schema_analysis, form_analysis, vba_modules, metrics)
    spec_md_path = db_dir / "app_spec.md"
    spec_md_path.write_text(md, encoding="utf-8")
    print(f"  Written: {spec_md_path}")

    # Print summary
    entity_counts: dict[str, int] = {}
    for e in schema_analysis["entities"].values():
        entity_counts[e["scope"]] = entity_counts.get(e["scope"], 0) + 1
    print(f"\n  Entities: {sum(entity_counts.values())} total")
    for scope, count in sorted(entity_counts.items()):
        print(f"    {scope}: {count}")
    class_counts: dict[str, int] = {}
    for m in vba_modules:
        class_counts[m["classification"]] = class_counts.get(m["classification"], 0) + 1
    print(f"  VBA modules: " + ", ".join(f"{k}={v}" for k, v in sorted(class_counts.items())))
    routed = sum(1 for f in form_analysis["forms"] if f.get("suggested_route"))
    mapped = sum(1 for f in form_analysis["forms"] if f.get("mapped_entity"))
    print(f"  React routes: {routed}  (forms mapped to an entity: {mapped})")


def main() -> None:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    output_root = repo_root / "migration_output"

    if args.output_dir:
        dirs = [Path(args.output_dir).resolve()]
    elif args.db_name:
        dirs = [(output_root / args.db_name).resolve()]
    else:
        dirs = find_db_dirs(output_root)
        if not dirs:
            print("No migration_output sub-folders with schema.json + forms_vba.json found.")
            return

    for db_dir in dirs:
        process_db(db_dir, repo_root, args.config)

    print("\nDone.")


if __name__ == "__main__":
    main()
