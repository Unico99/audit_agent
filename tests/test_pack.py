import io
import zipfile
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

from pack import sample_tb
from pack.api import app
from pack.engine import analyze_tb, fill_financials, fill_requirements

TPL = Path(__file__).resolve().parent.parent / "pack" / "templates"


@pytest.fixture(scope="module")
def tbs(tmp_path_factory):
    d = tmp_path_factory.mktemp("tb")
    sample_tb.build(d / "ok.xlsx")
    sample_tb.build(d / "bad.xlsx", {"1102": 585})            # فرق 585 = رصيد حساب واحد؟ (لا) — فرق عام
    sample_tb.build(d / "one_side.xlsx", {"7112": -2 * 61956})  # حساب اتسجل في الجهة الغلط
    return d


def test_balanced_ok(tbs):
    r = analyze_tb(tbs / "ok.xlsx")
    assert r["ok"] and r["balanced"] and not r["unmapped"]


def test_unbalanced_stops_with_reasons(tbs):
    r = analyze_tb(tbs / "bad.xlsx")
    assert not r["ok"] and r["difference"] == 585
    assert r["reasons"]


def test_hint_finds_one_sided_account(tbs):
    r = analyze_tb(tbs / "one_side.xlsx")
    assert abs(r["difference"]) == 2 * 61956
    assert any("7112" in h for h in r["hints"])


def test_unmapped_account_blocks(tbs, tmp_path):
    import pandas as pd
    df = pd.read_excel(tbs / "ok.xlsx")
    df.loc[len(df)] = ["9999", "حساب غريب", 0, 0, 0]
    df.to_excel(tmp_path / "u.xlsx", index=False)
    r = analyze_tb(tmp_path / "u.xlsx")
    assert r["unmapped"] and r["balanced"]


def test_fill_matches_known_figures(tbs, tmp_path):
    r = analyze_tb(tbs / "ok.xlsx")
    out = tmp_path / "fs.xlsx"
    res = fill_financials(r["df"], TPL / "financial_statements.xlsx", out, "شركة تجريبية", 2023)
    assert all(c[1] != "✘" for c in res["checks"])
    v = res["values"]
    assert v["c"]["assets"] == 28626893 and v["c"]["ni"] == -2552700
    assert v["p"]["ni"] == -3293242 and v["p"]["equity"] == -1870177
    wb = openpyxl.load_workbook(out)
    assert wb["قائمة المركز المالي"]["C17"].value == "=C13+C16"        # إجماليات حيّة
    assert "شركة تجريبية" in wb["قائمة المركز المالي"]["A1"].value
    assert wb["قائمة التدفقات النقدية"]["B14"].value == 775252
    assert wb["قائمة التدفقات النقدية"]["B21"].value == -615576
    assert wb["إيضاح 9 - ممتلكات ومعدات"]["E9"].value == "=SUM(B9:D9)"


def test_year_replaced(tbs, tmp_path):
    r = analyze_tb(tbs / "ok.xlsx")
    out = tmp_path / "fs.xlsx"
    fill_financials(r["df"], TPL / "financial_statements.xlsx", out, "ش", 2025)
    wb = openpyxl.load_workbook(out)
    txt = " ".join(str(c.value) for ws in wb for row in ws.iter_rows() for c in row if c.value)
    assert "2023" not in txt and "2022" not in txt and "2025" in txt


def test_requirements_year_and_refs(tmp_path):
    out = tmp_path / "rq.xlsx"
    fill_requirements(TPL / "audit_requirements.xlsx", out, "ش", 2026)
    wb = openpyxl.load_workbook(out)
    assert "2026" in wb["English"]["A1"].value and "2025" not in wb["English"]["A1"].value
    assert wb["English"]["E4"].value and wb["العربية"]["E3"].value
    assert wb["English"].max_row == 75 and wb["English"]["E75"].value is None


def test_api_flow(tbs):
    c = TestClient(app)
    bad = c.post("/check", files={"file": ("t.xlsx", (tbs / "bad.xlsx").read_bytes())}).json()
    assert not bad["ok"] and "الفرق" in bad["message"]
    r = c.post("/build", files={"file": ("t.xlsx", (tbs / "bad.xlsx").read_bytes())}, data={"client": "ش", "period": "2025"})
    assert r.status_code == 409
    ok = c.post("/build", files={"file": ("t.xlsx", (tbs / "ok.xlsx").read_bytes())}, data={"client": "ش", "period": "FY 2025"})
    assert ok.status_code == 200
    z = zipfile.ZipFile(io.BytesIO(ok.content))
    assert len(z.namelist()) == 2
    assert c.post("/build", files={"file": ("t.xlsx", b"x")}, data={"client": "ش", "period": "بدون سنة"}).status_code == 422
