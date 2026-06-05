"""
Artifact Inspector
Reads schema.json + forms_vba.json from a migration_output folder and produces:
  - app_spec.json  — compact machine-readable specification for code generation
  - app_spec.md    — human-readable overview with VBA classification and form→entity mapping

Usage:
    python generators/inspect_artifacts.py
    python generators/inspect_artifacts.py --db-name NorthwindStarterED
    python generators/inspect_artifacts.py --output-dir path/to/custom_dir
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

# ---------------------------------------------------------------------------
# VBA module classification rules
# ---------------------------------------------------------------------------

# Modules that contain portable business logic → C# Domain/Services
_DOMAIN_MODULES: set[str] = {
    "modOrders",
    "modInventory",
    "modPurchaseOrders",
    "modValidation",
    "modCompanies",
    "modMath",
    "modSecurity",
    "modStrings",
    "clsErrorHandler",
    "modTableDataMacros",
    "modReportParameters",
}

# Modules that are UI glue → translate to React state / hooks
_UI_GLUE_MODULES: set[str] = {
    "modForms",
    "modRibbonCallback",
    "modStartup",
    "modGlobal",
}

# Modules replaced by framework (EF Core, logging, file APIs)
_INFRA_MODULES: set[str] = {
    "modDAO",
    "modFiles",
    "modDebug",
}


def _classify_module(name: str, module_type: str) -> str:
    if name in _DOMAIN_MODULES:
        return "domain"
    if name in _UI_GLUE_MODULES:
        return "ui_glue"
    if name in _INFRA_MODULES:
        return "infrastructure"
    if module_type == "Document":
        # Form_* or Report_* — all UI glue
        return "ui_glue"
    if module_type == "ClassModule":
        return "domain"
    return "unknown"


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


def _to_pascal(name: str) -> str:
    parts = re.split(r"[_\s]+", name)
    return "".join(p[:1].upper() + p[1:] for p in parts if p)


# ---------------------------------------------------------------------------
# Core entity classification (v1 scope)
# ---------------------------------------------------------------------------

_V1_ENTITIES = {"Companies", "Contacts", "Products", "Orders", "OrderDetails"}
_V1_SUPPORT = {
    "CompanyTypes", "States", "TaxStatus", "Titles",
    "Employees", "Privileges", "EmployeePrivileges",
    "ProductCategories", "OrderStatus", "OrderDetailStatus",
}


def _entity_scope(table_name: str) -> str:
    if table_name in _V1_ENTITIES:
        return "v1_core"
    if table_name in _V1_SUPPORT:
        return "v1_support"
    return "later"


# ---------------------------------------------------------------------------
# Schema analysis
# ---------------------------------------------------------------------------


def analyse_schema(schema: dict) -> dict:
    entities = {}
    fk_graph: dict[str, list[dict]] = {}

    for table_name, table_data in schema.items():
        if not isinstance(table_data, dict):
            continue  # skip _queries, _linked_tables, etc.
        columns = table_data.get("columns", [])
        pks = set(table_data.get("primary_keys", []))
        fks = table_data.get("foreign_keys", [])
        row_count = table_data.get("row_count", 0)

        entity_name = _to_pascal(table_name)  # usually already PascalCase
        scope = _entity_scope(table_name)

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
# Form/Report analysis
# ---------------------------------------------------------------------------


def _form_to_entity(record_source: str | None, entities: dict) -> str | None:
    if not record_source:
        return None
    rs = record_source.strip().strip('"')
    # Direct table name match
    if rs in entities:
        return rs
    # Try without SELECT ... FROM prefix
    m = re.search(r"\bFROM\s+\[?(\w+)\]?", rs, re.IGNORECASE)
    if m and m.group(1) in entities:
        return m.group(1)
    # Partial match (record source may be a query name similar to table)
    for table in entities:
        if rs.lower() == table.lower():
            return table
    return None


def analyse_forms(forms_vba: dict, entities: dict) -> dict:
    forms_info = []
    for form in forms_vba.get("forms", []):
        name = form["name"]
        rs = form.get("record_source")
        entity = _form_to_entity(rs, entities)
        control_count = len(form.get("controls", []))
        event_count = sum(len(c.get("events", [])) for c in form.get("controls", []))
        is_subform = name.startswith("sfrm")
        is_dialog = "Dialog" in name or "Login" in name or "Credentials" in name
        forms_info.append({
            "name": name,
            "record_source": rs,
            "mapped_entity": entity,
            "control_count": control_count,
            "event_count": event_count,
            "is_subform": is_subform,
            "is_dialog": is_dialog,
            "suggested_route": _suggest_route(name, entity, is_subform),
        })

    reports_info = []
    for report in forms_vba.get("reports", []):
        name = report["name"]
        rs = report.get("record_source")
        entity = _form_to_entity(rs, entities)
        reports_info.append({
            "name": name,
            "record_source": rs,
            "mapped_entity": entity,
            "control_count": len(report.get("controls", [])),
            "suggested_component": _suggest_report_component(name),
        })

    return {"forms": forms_info, "reports": reports_info}


def _suggest_route(form_name: str, entity: str | None, is_subform: bool) -> str | None:
    if is_subform:
        return None  # rendered as child component, not a route

    # Map known forms
    route_map = {
        "frmCompanyList": "/companies",
        "frmCompanyDetail": "/companies/:id",
        "frmOrderList": "/orders",
        "frmOrderDetails": "/orders/:id",
        "frmProductList": "/products",
        "frmProductDetail": "/products/:id",
        "frmEmployeeList": "/employees",
        "frmPurchaseOrderList": "/purchase-orders",
        "frmPurchaseOrderDetails": "/purchase-orders/:id",
        "frmLogin": "/login",
        "frmAdmin": "/admin",
    }
    if form_name in route_map:
        return route_map[form_name]

    # Derive from entity
    if entity:
        slug = re.sub(r"(?<!^)(?=[A-Z])", "-", entity).lower()
        if "List" in form_name:
            return f"/{slug}"
        if "Detail" in form_name:
            return f"/{slug}/:id"
    return None


def _suggest_report_component(report_name: str) -> str:
    component_map = {
        "rptInvoice": "InvoiceReport",
        "rptSalesByProduct": "SalesByProductReport",
        "rptSalesByProductQuarterly": "SalesByProductQuarterlyReport",
        "rptSalesByEmployee": "SalesByEmployeeReport",
        "rptEmployeeEmailList": "EmployeeEmailListReport",
        "rptEmployeePhoneList": "EmployeePhoneListReport",
        "rptProductCatalog": "ProductCatalogReport",
    }
    return component_map.get(report_name, _to_pascal(report_name.lstrip("r").lstrip("pt")))


# ---------------------------------------------------------------------------
# VBA module analysis
# ---------------------------------------------------------------------------


def analyse_vba_modules(forms_vba: dict) -> list[dict]:
    modules = []
    for mod in forms_vba.get("vba_modules", []):
        name = mod["name"]
        mod_type = mod.get("type", "")
        classification = _classify_module(name, mod_type)
        modules.append({
            "name": name,
            "type": mod_type,
            "line_count": mod.get("line_count", 0),
            "classification": classification,
            "has_api_declarations": mod.get("has_api_declarations", False),
            "external_references": mod.get("external_references", []),
            "suggested_csharp_class": _suggest_csharp_class(name, classification),
        })
    # Sort: domain first, then ui_glue, infra, unknown
    order = {"domain": 0, "ui_glue": 1, "infrastructure": 2, "unknown": 3}
    modules.sort(key=lambda m: (order.get(m["classification"], 3), m["name"]))
    return modules


def _suggest_csharp_class(name: str, classification: str) -> str | None:
    if classification != "domain":
        return None
    mapping = {
        "modOrders": "OrderService",
        "modInventory": "InventoryService",
        "modPurchaseOrders": "PurchaseOrderService",
        "modValidation": "ValidationService",
        "modCompanies": "CompanyService",
        "modMath": "MathHelper",
        "modSecurity": "SecurityService",
        "modStrings": "StringHelper",
        "clsErrorHandler": "ErrorHandler",
        "modTableDataMacros": "TableDataService",
        "modReportParameters": "ReportParameterService",
    }
    return mapping.get(name, _to_pascal(name.lstrip("mod").lstrip("cls")) + "Service")


# ---------------------------------------------------------------------------
# Markdown generation
# ---------------------------------------------------------------------------


def build_markdown(schema_analysis: dict, form_analysis: dict, vba_modules: list[dict], metrics: dict) -> str:
    lines: list[str] = []
    lines.append("# App Specification — Northwind Modernization\n")
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

    # --- VBA port queue ---
    domain_mods = [m for m in vba_modules if m["classification"] == "domain"]
    if domain_mods:
        lines.append("## VBA Port Queue (recommended order)\n")
        queue = ["modValidation", "modCompanies", "modMath", "modStrings",
                 "modInventory", "modOrders", "modPurchaseOrders",
                 "modSecurity", "clsErrorHandler", "modTableDataMacros", "modReportParameters"]
        domain_names = {m["name"] for m in domain_mods}
        i = 1
        for name in queue:
            if name in domain_names:
                mod = next(m for m in domain_mods if m["name"] == name)
                lines.append(f"{i}. `{name}` ({mod['line_count']} lines) → `{mod['suggested_csharp_class']}`")
                i += 1
        # Any domain modules not in our queue order
        for mod in domain_mods:
            if mod["name"] not in queue:
                lines.append(f"{i}. `{mod['name']}` ({mod['line_count']} lines) → `{mod['suggested_csharp_class']}`")
                i += 1
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
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent
    default_output_root = repo_root / "migration_output"

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
    return parser.parse_args()


def find_db_dirs(output_root: Path) -> list[Path]:
    return sorted(
        p for p in output_root.iterdir()
        if p.is_dir() and (p / "schema.json").exists() and (p / "forms_vba.json").exists()
    )


def process_db(db_dir: Path) -> None:
    print(f"\n{'='*60}")
    print(f"Inspecting: {db_dir.name}")
    print(f"{'='*60}")

    schema_path = db_dir / "schema.json"
    forms_vba_path = db_dir / "forms_vba.json"

    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    forms_vba = json.loads(forms_vba_path.read_text(encoding="utf-8"))

    schema_analysis = analyse_schema(schema)
    form_analysis = analyse_forms(forms_vba, schema_analysis["entities"])
    vba_modules = analyse_vba_modules(forms_vba)
    metrics = forms_vba.get("metrics", {})

    app_spec = {
        "db_name": db_dir.name,
        "schema": schema_analysis,
        "forms": form_analysis["forms"],
        "reports": form_analysis["reports"],
        "vba_modules": vba_modules,
        "metrics": metrics,
    }

    spec_json_path = db_dir / "app_spec.json"
    spec_json_path.write_text(json.dumps(app_spec, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  Written: {spec_json_path}")

    md = build_markdown(schema_analysis, form_analysis, vba_modules, metrics)
    spec_md_path = db_dir / "app_spec.md"
    spec_md_path.write_text(md, encoding="utf-8")
    print(f"  Written: {spec_md_path}")

    # Print summary
    entity_counts = {}
    for e in schema_analysis["entities"].values():
        entity_counts[e["scope"]] = entity_counts.get(e["scope"], 0) + 1
    print(f"\n  Entities: {sum(entity_counts.values())} total")
    for scope, count in sorted(entity_counts.items()):
        print(f"    {scope}: {count}")
    domain_count = sum(1 for m in vba_modules if m["classification"] == "domain")
    print(f"  VBA modules to port: {domain_count}")
    routed = sum(1 for f in form_analysis["forms"] if f.get("suggested_route"))
    print(f"  React routes: {routed}")


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
        process_db(db_dir)

    print("\nDone.")


if __name__ == "__main__":
    main()
