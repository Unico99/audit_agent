import pandas as pd
import pytest

TB_COLS = ["كود الحساب", "اسم الحساب", "مدين", "دائن", "السنة السابقة"]

# ميزان متوازن (855,000)، صافي الربح 55,000
BALANCED = [
    ("1000", "النقدية بالصندوق", 20000, 0, 18000),
    ("1010", "البنك - الحساب الجاري", 80000, 0, 60000),
    ("1100", "العملاء", 120000, 0, 80000),
    ("1200", "المخزون", 90000, 0, 88000),
    ("1500", "الأصول الثابتة", 200000, 0, 200000),
    ("1590", "مجمع إهلاك الأصول الثابتة", 0, 40000, -20000),
    ("2000", "الموردين", 0, 70000, -65000),
    ("2100", "قروض بنكية", 0, 100000, -100000),
    ("2200", "مصروفات مستحقة", 0, 8000, -6000),
    ("3000", "رأس المال", 0, 200000, -200000),
    ("3100", "أرباح محتجزة", 0, 37000, -6000),
    ("4000", "إيرادات المبيعات", 0, 400000, -350000),
    ("5000", "تكلفة المبيعات", 250000, 0, 220000),
    ("5100", "مرتبات وأجور", 55000, 0, 40000),
    ("5200", "إيجار", 12000, 0, 12000),
    ("5300", "مصروف إهلاك", 20000, 0, 20000),
    ("5400", "فوائد بنكية", 8000, 0, 9000),
]


@pytest.fixture
def env(tmp_path, monkeypatch):
    """بيئة معزولة: فولدرات بيانات/مخرجات/نماذج/تخصيصات مؤقتة + ميزان متوازن وآخر غير متوازن."""
    dirs = {k: tmp_path / k for k in ("data", "out", "tpl", "map", "std")}
    for d in dirs.values():
        d.mkdir()
    pd.DataFrame(BALANCED, columns=TB_COLS).to_excel(dirs["data"] / "good.xlsx", index=False)
    bad = [list(r) for r in BALANCED]
    bad[0][2] += 5000  # فرق 5,000
    pd.DataFrame(bad, columns=TB_COLS).to_excel(dirs["data"] / "bad.xlsx", index=False)
    monkeypatch.setenv("AUDIT_DATA_DIR", str(dirs["data"]))
    monkeypatch.setenv("AUDIT_OUTPUT_DIR", str(dirs["out"]))
    monkeypatch.setenv("AUDIT_TEMPLATES_DIR", str(dirs["tpl"]))
    monkeypatch.setenv("AUDIT_MAPPINGS_DIR", str(dirs["map"]))
    monkeypatch.setenv("AUDIT_STANDARDS_DIR", str(dirs["std"]))
    return dirs
