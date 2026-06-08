"""Tests for the generic heuristics + override config in generators/inspect_artifacts.py."""

import pytest

import inspect_artifacts as ia


# --------------------------------------------------------------------------- #
# string helpers
# --------------------------------------------------------------------------- #


def test_to_pascal():
    assert ia._to_pascal("order_id") == "OrderId"
    assert ia._to_pascal("FirstName") == "FirstName"
    assert ia._to_pascal("company type") == "CompanyType"


@pytest.mark.xfail(reason="#8: ALL-CAPS / snake names aren't normalised before casing")
def test_to_pascal_all_caps_known_bug():
    assert ia._to_pascal("TOTAL_AMOUNT") == "TotalAmount"


def test_kebab():
    assert ia._kebab("PurchaseOrders") == "purchase-orders"
    assert ia._kebab("Companies") == "companies"


def test_csharp_type_nullability():
    assert ia._to_csharp_type("INT", True) == "int?"
    assert ia._to_csharp_type("INT", False) == "int"
    assert ia._to_csharp_type("NVARCHAR", True) == "string"   # ref types get no '?'
    assert ia._to_csharp_type("DATE", True) == "DateOnly?"
    assert ia._to_csharp_type("MYSTERY", True) == "object"


# --------------------------------------------------------------------------- #
# entity scope (structural)
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize("name,cols,out_fk,in_deg,expected", [
    ("States", 2, 0, 1, "v1_support"),          # leaf dimension referenced by others
    ("USysRibbons", 3, 0, 0, "later"),          # Access system table
    ("Welcome", 4, 0, 0, "later"),              # isolated small island
    ("Companies", 15, 3, 5, "v1_core"),         # connected business entity
    ("LinkTbl", 3, 2, 0, "v1_support"),         # small junction
    ("BigIsland", 20, 0, 0, "v1_core"),         # isolated but substantial
])
def test_classify_entity(name, cols, out_fk, in_deg, expected):
    assert ia._classify_entity(name, cols, out_fk, in_deg) == expected


# --------------------------------------------------------------------------- #
# VBA classification (content + convention)
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize("name,mtype,source,has_api,expected", [
    ("modOrders", "StandardModule", "Set rs = CurrentDb.OpenRecordset(...)", False, "domain"),
    ("modValidation", "StandardModule", "MsgBox \"bad\"", False, "domain"),
    ("clsErrorHandler", "ClassModule", "", False, "domain"),
    ("modForms", "StandardModule", "Forms!frmX.Requery", False, "ui_glue"),
    ("modGlobal", "StandardModule", "DoCmd.OpenForm \"x\"", False, "ui_glue"),
    ("modDAO", "StandardModule", "", False, "infrastructure"),       # name convention
    ("modFiles", "StandardModule", "", False, "infrastructure"),
    ("modAnything", "StandardModule", "Declare Function Foo Lib \"k\" ()", False, "infrastructure"),
    ("modAnything", "StandardModule", "", True, "infrastructure"),   # has_api flag
    ("Form_frmX", "Document", "", False, "ui_glue"),                 # code-behind
])
def test_classify_module(name, mtype, source, has_api, expected):
    assert ia._classify_module(name, mtype, source, has_api) == expected


def test_suggest_csharp_class():
    assert ia._suggest_csharp_class("modOrders", "domain") == "OrdersService"
    assert ia._suggest_csharp_class("modMath", "domain") == "MathService"
    assert ia._suggest_csharp_class("clsErrorHandler", "domain") == "ErrorHandler"  # ends in Handler
    assert ia._suggest_csharp_class("modForms", "ui_glue") is None


# --------------------------------------------------------------------------- #
# form/report → entity matching, routes, components
# --------------------------------------------------------------------------- #


def _entities(*names):
    return {n: {"table_name": n} for n in names}


def test_match_form_to_entity_singular_plural():
    ents = _entities("Companies", "Orders", "Products", "PurchaseOrders")
    assert ia._match_form_to_entity("frmCompanyList", ents) == "Companies"
    assert ia._match_form_to_entity("frmOrderDetails", ents) == "Orders"
    assert ia._match_form_to_entity("frmPurchaseOrderList", ents) == "PurchaseOrders"
    assert ia._match_form_to_entity("frmAbout", ents) is None


def test_suggest_route():
    assert ia._suggest_route("frmCompanyList", "Companies", False, False) == "/companies"
    assert ia._suggest_route("frmCompanyDetail", "Companies", False, False) == "/companies/:id"
    assert ia._suggest_route("sfrmX", "Companies", True, False) is None     # subform
    assert ia._suggest_route("frmLogin", None, False, True) is None         # dialog
    assert ia._suggest_route("frmThingList", None, False, False) == "/thing-list"
    assert ia._suggest_route("frmAbout", None, False, False) is None


def test_suggest_report_component():
    assert ia._suggest_report_component("rptInvoice") == "InvoiceReport"
    assert ia._suggest_report_component("srptGastronomic") == "GastronomicReport"
    assert ia._suggest_report_component("rptSalesByProduct") == "SalesByProductReport"


# --------------------------------------------------------------------------- #
# analyse_* integration + config overrides
# --------------------------------------------------------------------------- #


def _schema():
    return {
        "Lookup": {
            "columns": [{"name": "LookupID", "type_name": "LONG", "nullable": False},
                        {"name": "Label", "type_name": "NVARCHAR", "nullable": True}],
            "primary_keys": ["LookupID"],
            "foreign_keys": [],
            "row_count": 3,
        },
        "Main": {
            "columns": [{"name": "MainID", "type_name": "LONG", "nullable": False},
                        {"name": "LookupID", "type_name": "LONG", "nullable": True},
                        {"name": "Name", "type_name": "NVARCHAR", "nullable": True}],
            "primary_keys": ["MainID"],
            "foreign_keys": [{"fk_column": "LookupID", "pk_table": "Lookup", "pk_column": "LookupID"}],
            "row_count": 10,
        },
        "_queries": [],  # non-dict-ish metadata must be ignored
    }


def test_analyse_schema_default_scopes_and_fields():
    sa = ia.analyse_schema(_schema(), {})
    ents = sa["entities"]
    assert "_queries" not in ents
    assert ents["Lookup"]["scope"] == "v1_support"   # out_fk 0, referenced once
    assert ents["Main"]["scope"] == "v1_core"        # has an outgoing FK
    fld = {f["name"]: f for f in ents["Main"]["fields"]}
    assert fld["MainID"]["is_pk"] and fld["MainID"]["csharp_type"] == "int"
    assert fld["LookupID"]["is_fk"] and fld["LookupID"]["csharp_type"] == "int?"
    assert fld["Name"]["csharp_name"] == "Name"


def test_analyse_schema_scope_override():
    sa = ia.analyse_schema(_schema(), {"entity_scope": {"Main": "later", "Lookup": "v1_core"}})
    assert sa["entities"]["Main"]["scope"] == "later"
    assert sa["entities"]["Lookup"]["scope"] == "v1_core"


def test_analyse_vba_modules_with_overrides():
    fv = {"vba_modules": [
        {"name": "modOrders", "type": "StandardModule", "source": "CurrentDb", "line_count": 9,
         "has_api_declarations": False},
        {"name": "modDAO", "type": "StandardModule", "source": "", "line_count": 5,
         "has_api_declarations": False},
    ]}
    cfg = {"vba_classification": {"modDAO": "domain"}, "csharp_class": {"modOrders": "OrderService"}}
    mods = {m["name"]: m for m in ia.analyse_vba_modules(fv, cfg)}
    assert mods["modOrders"]["classification"] == "domain"
    assert mods["modOrders"]["suggested_csharp_class"] == "OrderService"   # overridden
    assert mods["modDAO"]["classification"] == "domain"                    # overridden from infra


def test_analyse_forms_name_mapping_and_route_override_and_null_suppression():
    ents = ia.analyse_schema(_schema(), {})["entities"]
    fv = {"forms": [
        {"name": "frmMainList", "record_source": None, "controls": []},
        {"name": "frmAdmin", "record_source": None, "controls": []},
        {"name": "frmMainBoard", "record_source": None, "controls": []},
    ], "reports": []}
    cfg = {"routes": {"frmAdmin": "/admin", "frmMainBoard": None}}
    forms = {f["name"]: f for f in ia.analyse_forms(fv, ents, cfg)["forms"]}
    # name match populates the entity even though record_source is null
    assert forms["frmMainList"]["mapped_entity"] == "Main"
    assert forms["frmMainList"]["suggested_route"] == "/main"
    assert forms["frmAdmin"]["suggested_route"] == "/admin"        # override applied
    assert forms["frmMainBoard"]["suggested_route"] is None        # null-suppressed


def test_fk_topological_sort_dependencies_first():
    ents = ia.analyse_schema(_schema(), {})["entities"]
    order = ia._fk_topological_sort(ents)
    assert order.index("Lookup") < order.index("Main")
