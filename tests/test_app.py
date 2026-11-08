from pathlib import Path

import pytest

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest  # noqa: E402

from agent import fs  # noqa: E402

APP = str(Path(__file__).resolve().parent.parent / "app.py")


def _run():
    return AppTest.from_file(APP, default_timeout=60).run()


def _select(at, name):
    at.selectbox(key="selected_file").set_value(name).run()
    assert not at.exception


def test_unbalanced_file_locks_gated_buttons(env):
    at = _run()                       # الافتراضي bad.xlsx (أول ملف أبجديًا)
    assert not at.exception
    assert at.button(key="quick_req").disabled and at.button(key="quick_fs").disabled
    assert any("غير متوازن" in e.value for e in at.sidebar.error)


def test_balanced_file_unlocks_buttons(env):
    at = _run()
    _select(at, "good.xlsx")
    assert not at.button(key="quick_req").disabled and not at.button(key="quick_fs").disabled
    assert any("متوازن" in s.value for s in at.sidebar.success)


def test_quick_workpaper_then_approve_writes_file(env):
    at = _run()                       # bad.xlsx: ورقة العمل مش متقفلة بالتوازن
    at.button(key="quick").click().run()
    assert not at.exception and at.session_state["pending"]["name"] == "write_workpaper"
    assert not list(env["out"].glob("*.xlsx"))
    at.button(key="approve").click().run()
    assert not at.exception and at.session_state["saved_file"]
    assert len(list(env["out"].glob("*.xlsx"))) == 1


def test_reject_writes_nothing(env):
    at = _run()
    at.button(key="quick").click().run()
    at.button(key="reject").click().run()
    assert at.session_state["pending"] is None and not list(env["out"].glob("*.xlsx"))


def test_requirements_flow_on_balanced_file(env):
    at = _run()
    _select(at, "good.xlsx")
    at.text_input(key="client_name").set_value("شركة النيل").run()
    at.text_input(key="period").set_value("2025-12-31").run()
    at.button(key="quick_req").click().run()
    assert not at.exception
    assert at.session_state["pending"]["name"] == "write_requirements"
    assert not list(env["out"].glob("*.xlsx"))              # لسه ما اتكتبش قبل الموافقة
    at.button(key="approve").click().run()
    assert not at.exception
    files = list(env["out"].glob("متطلبات*.xlsx"))
    assert len(files) == 1 and "شركة_النيل" in files[0].name


def test_fs_tab_approve_mapping_then_build(env):
    at = _run()
    _select(at, "good.xlsx")
    at.text_input(key="client_name").set_value("شركة النيل").run()
    assert at.button(key="fs_build").disabled                # قبل اعتماد التخصيص: مقفول
    at.button(key="fs_save").click().run()                   # اعتماد الاقتراح الآلي كما هو
    assert not at.exception
    saved = fs.load_mapping(fs.mapping_key("شركة النيل"))
    assert saved and saved["approved"]
    at.button(key="fs_build").click().run()
    assert not at.exception
    r = at.session_state["fs_result"]
    assert r["ok"] and r["all_checks_passed"] and r["net_income"] == 55000
    assert len(list(env["out"].glob("القوائم_المالية*.xlsx"))) == 1


def test_fs_tab_blocked_when_unbalanced(env):
    at = _run()                                              # bad.xlsx
    assert not at.exception
    assert any("متاحة بعد التأكد" in e.value for e in at.error)


def test_templates_tab_renders_and_saves_config(env):
    import io

    from agent import sample_templates, templates

    for kind, make in (("requirements", sample_templates.make_requirements_template),
                       ("financial_statements", sample_templates.make_fs_template)):
        tmp = env["tpl"] / "_x.xlsx"
        make(tmp)
        templates.register(kind, tmp.read_bytes(), "mine.xlsx", name="mine")
    at = _run()
    assert not at.exception
    at.button(key="save_cfg_requirements").click().run()
    assert not at.exception
    at.button(key="save_cfg_financial_statements").click().run()
    assert not at.exception


def test_real_mode_without_key_is_blocked(env, monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    at = _run()
    at.radio(key="mode_choice").set_value("كلود الحقيقي").run()
    at.button(key="quick").click().run()
    assert not at.exception and at.error
