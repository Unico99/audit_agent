"""تصنيف الحسابات (مجموعة + فئة) وبناء جدول الحسابات بأرصدتها.

التصنيف بالكلمات المفتاحية في اسم الحساب أولًا، ثم بأول رقم في الكود
(1 أصول، 2 التزامات، 3 حقوق ملكية، 4 إيرادات، 5-9 مصروفات).
ده اقتراح بس؛ المراجع يعدّل المجموعة في شاشة التخصيص.
"""
from __future__ import annotations

import pandas as pd

from .textutil import norm

GROUP_SIDE = {
    "asset": "debit",
    "expense": "debit",
    "liability": "credit",
    "equity": "credit",
    "revenue": "credit",
}
GROUP_LABELS = {
    "asset": "أصل",
    "liability": "التزام",
    "equity": "حقوق ملكية",
    "revenue": "إيراد",
    "expense": "مصروف",
}
BS_GROUPS = {"asset", "liability", "equity"}
IS_GROUPS = {"revenue", "expense"}

# الترتيب مهم: أول قاعدة تتطابق هي اللي تتطبق
_RULES = [
    ("accum_dep", "asset", ["مجمع اهلاك", "مجمع الاهلاك", "مجمع استهلاك", "accumulated dep"]),
    ("accruals", "liability", ["مستحق", "accrued", "ارصده دائنه"]),
    ("cogs", "expense", ["تكلفه المبيعات", "تكلفه البضاعه", "تكلفه الايراد", "cost of sales", "cost of goods", "cogs"]),
    ("interest", "expense", ["فوائد", "فائده", "مصروفات تمويليه", "مصروف تمويلي", "interest"]),
    ("loans", "liability", ["قرض", "قروض", "تمويل بنكي", "سلفه بنكيه", "loan", "borrowing", "overdraft", "مكشوف"]),
    ("depreciation_exp", "expense", ["اهلاك", "استهلاك", "depreciation"]),
    ("payroll", "expense", ["مرتب", "اجور", "رواتب", "salar", "wages", "payroll"]),
    ("rent", "expense", ["ايجار", "rent"]),
    ("revenue", "revenue", ["ايراد", "مبيعات", "sales", "revenue"]),
    ("receivables", "asset", ["عملاء", "مدينون", "اوراق قبض", "ذمم مدينه", "receivable", "customers", "debtors"]),
    ("inventory", "asset", ["مخزون", "بضاعه", "inventory", "stock"]),
    ("fixed_assets", "asset", ["اصول ثابته", "الاصول الثابته", "اصل ثابت", "معدات", "الات", "سيارات", "مباني", "اراضي", "اثاث", "عقارات", "fixed asset", "property", "equipment", "vehicle", "furniture", "building"]),
    ("payables", "liability", ["مورد", "دائنون", "اوراق دفع", "ذمم دائنه", "payable", "suppliers", "creditors"]),
    ("equity_capital", "equity", ["راس المال", "capital"]),
    ("equity_retained", "equity", ["محتجز", "ارباح مرحله", "retained", "احتياطي", "reserve"]),
    ("cash", "asset", ["نقد", "خزين", "صندوق", "بنك", "بنوك", "cash", "bank"]),
    ("tax", None, ["ضريب", "tax", "vat", "قيمه مضافه"]),
]
_DIGIT_GROUP = {"1": "asset", "2": "liability", "3": "equity", "4": "revenue"}
_RULES = [(c, g, [norm(k) for k in kws]) for c, g, kws in _RULES]


def code_str(v) -> str:
    """كود الحساب كنص نظيف (1000.0 -> 1000)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _group_from_code(code: str):
    if not code:
        return None
    d = code[0]
    if d in _DIGIT_GROUP:
        return _DIGIT_GROUP[d]
    if d in "56789":
        return "expense"
    return None


def classify_account(code: str, name: str) -> dict:
    n = norm(name)
    for cat, group, kws in _RULES:
        if any(k in n for k in kws):
            g = group or _group_from_code(code) or "expense"
            return {"category": cat, "group": g}
    g = _group_from_code(code)
    return {"category": f"other_{g}" if g else "unknown", "group": g}


def accounts_table(df: pd.DataFrame, group_overrides: dict | None = None) -> pd.DataFrame:
    """جدول الحسابات: المجموعة والجانب الطبيعي والرصيد الحالي والسابق بإشارة طبيعية."""
    d = pd.to_numeric(df["debit"], errors="coerce").fillna(0.0)
    c = pd.to_numeric(df["credit"], errors="coerce").fillna(0.0)
    has_prior = "prior_balance" in df.columns
    prior = pd.to_numeric(df["prior_balance"], errors="coerce") if has_prior else None

    rows = []
    for i in df.index:
        code = code_str(df.at[i, "account_code"])
        nm = df.at[i, "account_name"]
        name = "" if pd.isna(nm) else str(nm).strip()
        cls = classify_account(code, name)
        group = (group_overrides or {}).get(code) or cls["group"]
        side = GROUP_SIDE.get(group)
        net = float(d[i] - c[i])
        amount = None if side is None else (net if side == "debit" else -net)
        pr = None
        if has_prior and side and not pd.isna(prior[i]):
            pr = float(prior[i]) if side == "debit" else -float(prior[i])
        rows.append(
            {
                "excel_row": int(i) + 2,
                "code": code,
                "name": name,
                "category": cls["category"],
                "group": group,
                "side": side,
                "net_debit": net,
                "amount": amount,
                "prior": pr,
            }
        )
    return pd.DataFrame(rows)


def net_income(acc: pd.DataFrame, col: str = "amount") -> float:
    rev = acc.loc[acc["group"] == "revenue", col].sum(skipna=True)
    exp = acc.loc[acc["group"] == "expense", col].sum(skipna=True)
    return float(rev - exp)
