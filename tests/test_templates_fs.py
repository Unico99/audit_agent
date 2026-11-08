import io

import pytest
from openpyxl import Workbook, load_workbook

from agent import checks, fs, requirements, templates
from agent.classify import accounts_table, classify_account
from agent.tools import run_tool


# ------------------------------------------------------------ تصنيف
@pytest.mark.parametrize(
    "code,name,cat,group",
    [
        ("1000", "النقدية بالصندوق", "cash", "asset"),
        ("1500", "الأصول الثابتة", "fixed_assets", "asset"),
        ("1590", "مجمع إهلاك الأصول الثابتة", "accum_dep", "asset"),
        ("2100", "قروض بنكية", "loans", "liability"),
        ("2200", "مصروفات مستحقة", "accruals", "liability"),
        ("4000", "إيرادات المبيعات", "revenue", "revenue"),
        ("5000", "تكلفة المبيعات", "cogs", "expense"),
        ("5400", "فوائد القروض", "interest", "expense"),
        ("5300", "مصروف إهلاك", "depreciation_exp", "expense"),
        ("2500", "حساب غريب", "other_liability", "liability"),
        ("", "بدون كود", "unknown", None),
    ],
)
def test_classify(code, name, cat, group):
    r = classify_account(code, name)
    assert (r["category"], r["group"]) == (cat, group)


# ------------------------------------------------------------ بوابة التوازن (في الكود)
@pytest.mark.parametrize("tool,extra", [
    ("get_requirements_context", {}),
    ("write_requirements", {"items": [{"area": "a", "requirement": "r", "priority": "low"}]}),
    ("build_financial_statements", {}),
])
def test_gate_blocks_unbalanced(env, tool, extra):
    out, is_err = run_tool(tool, {"path": "bad.xlsx", **extra})
    assert is_err and "غير متوازن" in out["error"]
    assert not list(env["out"].glob("*.xlsx"))


# ------------------------------------------------------------ قالب المتطلبات
def _custom_req_template(headers, header_row=3, sample_rows=0):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "شركة {{client}} - {{period}}"
    for i, h in enumerate(headers, start=1):
        ws.cell(header_row, i, h)
    for r in range(sample_rows):
        ws.cell(header_row + 1 + r, 1, "مثال")
    ws.cell(header_row + 1 + sample_rows + 2, 1, "توقيع المراجع")  # فوتر تحت الجدول
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def test_detect_requirements_columns_arabic_and_english():
    import tempfile, pathlib
    data = _custom_req_template(["Serial", "البند", "الإجراء المطلوب", "Documents", "المعيار", "Priority"])
    p = pathlib.Path(tempfile.mkdtemp()) / "t.xlsx"
    p.write_bytes(data)
    cfg = templates.detect_requirements_config(p)
    assert cfg["header_row"] == 3
    fields = set(cfg["columns"].values())
    assert {"area", "requirement", "evidence", "standard_reference", "priority"} <= fields


def test_register_once_and_reuse_with_custom_template(env):
    meta = templates.register("requirements", _custom_req_template(["م", "المجال", "المتطلب", "المرجع"], sample_rows=2), "mine.xlsx", name="قالب المكتب")
    assert templates.active_name("requirements") == meta["name"]
    items = [{"area": "نقدية", "requirement": f"متطلب {i}", "evidence": "x", "standard_reference": "ISA 505 (مقترح)", "priority": "high"} for i in range(5)]
    res = requirements.write_requirements(items, "شركة س", "2025", env["out"])
    ws = load_workbook(env["out"] / res["saved"]).active
    assert ws["A1"].value == "شركة شركة س - 2025"              # placeholders
    assert [ws.cell(4 + i, 3).value for i in range(5)] == [f"متطلب {i}" for i in range(5)]   # أمثلة القالب اتمسحت
    assert ws.cell(4, 1).value == 1 and ws.cell(5, 1).value == 2   # ترقيم تلقائي
    assert ws.cell(4, 4).value == "ISA 505 (مقترح)"
    labels = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
    assert "توقيع المراجع" in labels                              # الفوتر ما اتمسحش
    assert labels.index("توقيع المراجع") > 8                      # واتنزل تحت الجدول الأطول


def test_template_survives_new_session(env):
    templates.register("requirements", _custom_req_template(["م", "المتطلب"]), "m.xlsx", name="x")
    # "جلسة جديدة": نفس الفولدر، من غير أي حالة في الذاكرة
    path, meta = templates.get_active("requirements")
    assert meta["name"] == "x" and path.exists()


def test_register_rejects_non_xlsx(env):
    with pytest.raises(ValueError):
        templates.register("requirements", b"not excel", "x.xlsx")
    with pytest.raises(ValueError):
        templates.register("requirements", b"", "x.xls")


def test_requirements_context_uses_balances(env):
    out, err = run_tool("get_requirements_context", {"path": "good.xlsx"})
    assert not err
    cats = {c["category"] for c in out["categories_present"]}
    assert {"cash", "receivables", "inventory", "fixed_assets", "payables", "loans", "revenue", "payroll"} <= cats
    assert any("120,000" in i["requirement"] for i in out["suggested_items"])     # {balance} اتعبّى
    assert all("(مقترح)" in i["standard_reference"] for i in out["suggested_items"])


# ------------------------------------------------------------ القوائم المالية
def _approved_mapping(env, over=None):
    path, meta = templates.get_active("financial_statements")
    df = checks.load_trial_balance(env["data"] / "good.xlsx")
    lines = fs.template_lines(path, meta["config"])
    sug = fs.suggest_mapping(accounts_table(df), lines)
    m = {"template": meta["name"], "approved": True, "labels": {l["ref"]: l["label"] for l in lines},
         "accounts": {c: {"line": v["line"]} for c, v in sug.items()}}
    m["accounts"][fs.NI_KEY] = {"line": fs.suggest_ni_line(lines)}
    for c, ref in (over or {}).items():
        m["accounts"][c] = {"line": ref}
    return m, df, path, meta


def test_fs_build_ties_out(env):
    m, df, path, meta = _approved_mapping(env)
    assert all(v["line"] for v in m["accounts"].values())          # الاقتراح الآلي غطّى كل حاجة
    r = fs.build(df, m, path, meta["config"], env["out"], "شركة", "2025")
    assert r["ok"] and r["all_checks_passed"] and r["net_income"] == 55000
    ws = load_workbook(env["out"] / r["saved"])["قائمة المركز المالي"]
    vals = {ws.cell(i, 2).value: ws.cell(i, 4).value for i in range(6, 30)}
    assert vals["النقدية وما في حكمها"] == 100000                # 20,000 + 80,000
    assert vals["الأصول الثابتة (بالصافي)"] == 160000            # 200,000 - 40,000 مجمع الإهلاك
    assert vals["أرباح (خسائر) العام"] == 55000
    assert ws["D15"].value.startswith("=")                         # الإجماليات فضلت معادلات النموذج


def test_fs_wrong_mapping_is_caught(env):
    m, df, path, meta = _approved_mapping(env, {"1100": "قائمة المركز المالي!26"})   # عملاء -> موردين
    r = fs.build(df, m, path, meta["config"], env["out"])
    assert r["ok"] and not r["all_checks_passed"]
    assert any(c["status"] == "fail" for c in r["checks"])


def test_fs_blocks_incomplete_mapping(env):
    m, df, path, meta = _approved_mapping(env)
    m["accounts"]["1200"] = {"line": None}
    assert "غير مخصصة" in " ".join(fs.build(df, m, path, meta["config"], env["out"])["errors"])
    m, *_ = _approved_mapping(env, {"1200": "قائمة المركز المالي!15"})                    # صف إجمالي
    assert "إجمالي" in " ".join(fs.build(df, m, path, meta["config"], env["out"])["errors"])
    m, *_ = _approved_mapping(env)
    m["accounts"].pop(fs.NI_KEY)
    assert "أرباح" in " ".join(fs.build(df, m, path, meta["config"], env["out"])["errors"])


def test_fs_detects_template_change(env):
    m, df, path, meta = _approved_mapping(env)
    m["labels"]["قائمة المركز المالي!13"] = "بند قديم"
    assert "اتغيّر" in " ".join(fs.build(df, m, path, meta["config"], env["out"])["errors"])


def test_fs_duplicate_codes_blocked(env):
    m, df, path, meta = _approved_mapping(env)
    df.loc[1, "account_code"] = df.loc[0, "account_code"]
    assert "مكررة" in " ".join(fs.build(df, m, path, meta["config"], env["out"])["errors"])


def test_fs_flags_abnormal_balance(env):
    m, df, path, meta = _approved_mapping(env)
    df.loc[df["account_code"].astype(str) == "1100", ["debit", "credit"]] = [0, 120000]   # عملاء برصيد دائن
    # نعوّض في حساب آخر عشان الميزان يفضل متوازن
    df.loc[df["account_code"].astype(str) == "1200", "debit"] = 90000 + 240000
    r = fs.build(df, m, path, meta["config"], env["out"])
    assert any(a["code"] == "1100" for a in r["abnormal_balances"])


def test_unbalanced_fs_sheet_names_other_statement(env):
    """ورقة تدفقات نقدية في النموذج بتتساب وبتتبلّغ في التحذيرات."""
    path, meta = templates.get_active("financial_statements")
    wb = load_workbook(path)
    ws = wb.create_sheet("قائمة التدفقات النقدية")
    ws["B6"] = "التدفقات التشغيلية"
    b = io.BytesIO()
    wb.save(b)
    meta2 = templates.register("financial_statements", b.getvalue(), "mine.xlsx", name="mine")
    assert meta2["config"]["sheets"]["قائمة التدفقات النقدية"]["statement"] == "other"
    m, df, p2, m2 = _approved_mapping(env)
    r = fs.build(df, m, p2, m2["config"], env["out"])
    assert r["ok"] and any("التدفقات" in w for w in r["warnings"])


def test_build_tool_requires_approved_mapping(env):
    out, err = run_tool("build_financial_statements", {"path": "good.xlsx", "client": "ع"})
    assert err and "تخصيص معتمد" in out["error"]
    m, *_ = _approved_mapping(env)
    fs.save_mapping(fs.mapping_key("ع"), m)
    out, err = run_tool("build_financial_statements", {"path": "good.xlsx", "client": "ع", "period": "2025"})
    assert not err and out["ok"] and out["all_checks_passed"]


def test_formula_evaluator_matches_libreoffice_semantics():
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["A2"], ws["A3"] = 10, 20, 5
    ws["B1"] = "=SUM(A1:A3)"
    ws["B2"] = "=A1+A2-A3"
    ws["B3"] = "=B1-SUM(A1,A3)*2"
    ws["B4"] = "=IF(A1>1,1,0)"
    ev = fs.Evaluator(wb)
    assert ev.cell("S", "B1") == 35 and ev.cell("S", "B2") == 25 and ev.cell("S", "B3") == 5
    with pytest.raises(fs.NotSupported):
        ev.cell("S", "B4")


def _english_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Balance Sheet"
    ws["A4"], ws["C4"], ws["D4"] = "Description", 2025, 2024
    rows = [("Cash and cash equivalents", None), ("Trade receivables", None), ("Inventories", None),
            ("Property, plant and equipment", None), ("Total assets", "=SUM({c}5:{c}8)"),
            ("Share capital", None), ("Retained earnings", None), ("Profit for the year", None),
            ("Bank loans", None), ("Trade payables", None), ("Accrued expenses", None),
            ("Total equity and liabilities", "=SUM({c}10:{c}15)")]
    for i, (lab, f) in enumerate(rows, start=5):
        ws.cell(i, 1, lab)
        if f:
            ws.cell(i, 3, f.format(c="C"))
            ws.cell(i, 4, f.format(c="D"))
    w2 = wb.create_sheet("Income Statement")
    w2["A4"], w2["C4"], w2["D4"] = "Description", 2025, 2024
    for i, (lab, f) in enumerate([("Revenue", None), ("Cost of sales", None), ("Gross profit", "={c}5-{c}6"),
                                  ("Salaries and wages", None), ("Rent", None), ("Depreciation", None),
                                  ("Finance costs", None), ("Net profit for the year", "={c}7-SUM({c}8:{c}11)")], start=5):
        w2.cell(i, 1, lab)
        if f:
            w2.cell(i, 3, f.format(c="C"))
            w2.cell(i, 4, f.format(c="D"))
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def test_english_template_maps_and_is_actually_verified(env):
    """نموذج إنجليزي بتخطيط مختلف: لازم يتخصص آليًا وفحص الإجماليات يتعمل فعلاً (مش skipped)."""
    meta = templates.register("financial_statements", _english_template(), "acme.xlsx", name="acme")
    assert meta["config"]["sheets"]["Balance Sheet"]["current_col"] == "C"      # من عناوين السنوات 2025/2024
    m, df, path, meta = _approved_mapping(env)
    assert all(v["line"] for v in m["accounts"].values())
    r = fs.build(df, m, path, meta["config"], env["out"], "ACME", "2025")
    assert r["ok"] and r["all_checks_passed"], r["checks"]
    assert {c["status"] for c in r["checks"]} == {"ok"} and len(r["checks"]) == 4
    # ولو التخصيص غلط، النموذج الإنجليزي كمان يكشفه
    bad, *_ = _approved_mapping(env, {"1100": "Balance Sheet!14"})
    r2 = fs.build(df, bad, path, meta["config"], env["out"])
    assert not r2["all_checks_passed"] and any(c["status"] == "fail" for c in r2["checks"])


def test_all_checks_passed_is_false_when_checks_skipped(env):
    """لو القالب مفيهوش صفوف إجماليات معروفة، متقولش إن الفحص نجح."""
    wb = Workbook()
    ws = wb.active
    ws.title = "قائمة المركز المالي"
    ws["B5"], ws["D5"] = "البيان", "السنة الحالية"
    for i, lab in enumerate(["النقدية", "العملاء", "المخزون", "الأصول الثابتة", "رأس المال", "الأرباح المحتجزة", "أرباح العام", "قروض", "الموردين", "مستحقات"], start=6):
        ws.cell(i, 2, lab)
    w2 = wb.create_sheet("قائمة الدخل")
    w2["B5"], w2["D5"] = "البيان", "السنة الحالية"
    for i, lab in enumerate(["إيرادات المبيعات", "تكلفة المبيعات", "مرتبات", "إيجار", "إهلاك", "مصروفات تمويلية"], start=6):
        w2.cell(i, 2, lab)          # من غير أي صف إجمالي أو صافي ربح
    b = io.BytesIO()
    wb.save(b)
    templates.register("financial_statements", b.getvalue(), "plain.xlsx", name="plain")
    m, df, path, meta = _approved_mapping(env)
    r = fs.build(df, m, path, meta["config"], env["out"])
    assert r["ok"] and not r["all_checks_passed"]
    assert all(c["status"] == "skipped" for c in r["checks"])
