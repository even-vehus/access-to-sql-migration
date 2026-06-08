"""Tests for extract/extract_forms_vba.py — scoped process kill (#4), SaveAsText
parsing, encoding, and complexity metrics. (No live Access/COM is launched.)"""

import time

import pytest

import extract_forms_vba as e


# --------------------------------------------------------------------------- #
# #4 — PID resolution + scoped timeout kill
# --------------------------------------------------------------------------- #


class _BoomApp:
    @property
    def hWndAccessApp(self):
        raise RuntimeError("hung")


class _ZeroApp:
    hWndAccessApp = 0


def test_get_access_pid_is_robust_to_bad_objects():
    assert e._get_access_pid(_BoomApp()) is None
    assert e._get_access_pid(_ZeroApp()) is None


def test_run_with_timeout_fast_path_returns_value():
    assert e._run_with_timeout(lambda: 42, timeout=5) == 42


def _patch_taskkill(monkeypatch):
    calls = []

    class _R:
        returncode = 0

    monkeypatch.setattr(e.subprocess, "run", lambda *a, **k: (calls.append(list(a[0])), _R())[1])
    return calls


def test_timeout_kills_only_our_pid_never_image(monkeypatch):
    calls = _patch_taskkill(monkeypatch)
    monkeypatch.setattr(e, "_access_pid", 987654)
    with pytest.raises(e.COMTimeoutError):
        e._run_with_timeout(lambda: time.sleep(30), timeout=1)
    assert calls, "expected a taskkill invocation"
    assert calls[-1][:4] == ["taskkill", "/F", "/T", "/PID"]
    assert calls[-1][4] == "987654"
    assert all("/IM" not in tok for tok in calls[-1]), "must never use a broad image kill"


def test_timeout_with_unknown_pid_does_not_kill_anything(monkeypatch):
    calls = _patch_taskkill(monkeypatch)
    monkeypatch.setattr(e, "_access_pid", None)
    with pytest.raises(e.COMTimeoutError):
        e._run_with_timeout(lambda: time.sleep(30), timeout=1)
    assert calls == [], "with no known PID it must not call taskkill at all"


# --------------------------------------------------------------------------- #
# encoding-aware file read
# --------------------------------------------------------------------------- #


def test_read_access_text_file_handles_utf16_and_utf8(tmp_path):
    u16 = tmp_path / "u16.txt"
    u16.write_bytes(b"\xff\xfe" + "héllo".encode("utf-16-le"))
    assert e._read_access_text_file(u16) == "héllo"

    u8 = tmp_path / "u8.txt"
    u8.write_bytes("plain".encode("utf-8"))
    assert e._read_access_text_file(u8) == "plain"


# --------------------------------------------------------------------------- #
# SaveAsText parsing
# --------------------------------------------------------------------------- #


_FORM_TEXT = """\
Begin Form
    RecordSource ="Companies"
    Begin TextBox
        Name ="txtName"
        ControlSource ="CompanyName"
        AfterUpdate ="[Event Procedure]"
    End
    Begin CommandButton
        Name ="cmdSave"
        OnClick ="[Event Procedure]"
    End
End
"""


def test_parse_form_text_extracts_record_source_controls_and_events():
    parsed = e._parse_form_text(_FORM_TEXT, "frmCompany")
    assert parsed["name"] == "frmCompany"
    assert parsed["record_source"] == "Companies"
    ctrls = {c["name"]: c for c in parsed["controls"]}
    assert set(ctrls) == {"txtName", "cmdSave"}
    assert ctrls["txtName"]["control_source"] == "CompanyName"
    assert ctrls["txtName"]["events"] == ["AfterUpdate"]
    assert ctrls["cmdSave"]["events"] == ["OnClick"]


def test_parse_form_text_empty_returns_none():
    assert e._parse_form_text("   \n  ", "x") is None


# --------------------------------------------------------------------------- #
# complexity metrics
# --------------------------------------------------------------------------- #


def test_compute_metrics_scoring_and_tier():
    forms = [{"controls": [
        {"type": "Subform", "events": []},
        {"type": "ComboBox", "events": ["OnClick"]},
    ]}]
    vba = [{"line_count": 600, "has_api_declarations": True, "external_references": ["Shell"]}]
    metrics = e.compute_metrics(forms, [], vba, [])
    assert metrics["vba_total_lines"] == 600
    assert metrics["has_api_declarations"] is True
    assert metrics["has_external_objects"] is True
    assert metrics["subform_count"] == 1
    assert metrics["combo_box_count"] == 1
    # 600 lines(+3) + subform(+1) + api(+3) + external(+2) = 9 → complex
    assert metrics["complexity_score"] == 9
    assert metrics["complexity_tier"] == "complex"
