"""فحوصات حتمية (deterministic) بكود Python.

كل الحسابات والمطابقات هنا، مش في كلود. كلود بيستدعي الدوال دي كأدوات
وبيفسّر نتائجها بس.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# أسماء بديلة للأعمدة (إنجليزي / عربي) عشان الملفات الحقيقية مش دايمًا بنفس الشكل
COLUMN_ALIASES = {
    "account_code": ["account_code", "code", "account no", "كود الحساب", "رقم الحساب", "كود"],
    "account_name": ["account_name", "name", "account", "اسم الحساب", "البيان"],
    "debit": ["debit", "dr", "مدين"],
    "credit": ["credit", "cr", "دائن"],
    "prior_balance": ["prior_balance", "prior", "prior_year", "السنة السابقة", "رصيد سابق"],
}
REQUIRED = ("account_code", "account_name", "debit", "credit")


def load_trial_balance(path: str | Path, sheet: int | str = 0) -> pd.DataFrame:
    """بيقرا ميزان المراجعة (xlsx/xls/csv) ويوحّد أسماء الأعمدة."""
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, sheet_name=sheet)
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        raise ValueError(f"نوع ملف غير مدعوم: {path.suffix}")

    rename = {}
    for target, aliases in COLUMN_ALIASES.items():
        lowered = {a.lower() for a in aliases}
        for col in df.columns:
            if str(col).strip().lower() in lowered:
                rename[col] = target
                break
    df = df.rename(columns=rename)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            f"أعمدة ناقصة: {missing}. الأعمدة الموجودة: {list(df.columns)}"
        )
    return df


def _num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def check_balance(df: pd.DataFrame, tolerance: float = 0.01) -> dict:
    """هل مجموع المدين = مجموع الدائن؟"""
    total_debit = round(float(_num(df["debit"]).fillna(0).sum()), 2)
    total_credit = round(float(_num(df["credit"]).fillna(0).sum()), 2)
    diff = round(total_debit - total_credit, 2)
    return {
        "check": "trial_balance_balances",
        "total_debit": total_debit,
        "total_credit": total_credit,
        "difference": diff,
        "balanced": abs(diff) <= tolerance,
    }


def check_data_quality(df: pd.DataFrame) -> list[dict]:
    """جودة البيانات: حقول فاضية، أكواد مكررة، قيم مش رقمية، سالب، مدين ودائن مع بعض."""
    issues: list[dict] = []

    def add(kind: str, rows, detail: str) -> None:
        rows = list(rows)
        if rows:
            # رقم الصف في إكسل = index + 2 (الهيدر + العد من 1)
            issues.append(
                {"issue": kind, "excel_rows": [int(r) + 2 for r in rows], "detail": detail}
            )

    add("missing_account_code", df.index[df["account_code"].isna()], "كود حساب فاضي")
    add("missing_account_name", df.index[df["account_name"].isna()], "اسم حساب فاضي")

    dup = df[df["account_code"].notna() & df["account_code"].duplicated(keep=False)]
    if not dup.empty:
        codes = sorted(dup["account_code"].astype(str).unique())
        add("duplicate_account_code", dup.index, f"أكواد مكررة: {codes}")

    for col in ("debit", "credit"):
        raw = df[col]
        bad = raw.notna() & _num(raw).isna()
        add(f"non_numeric_{col}", df.index[bad], f"قيم مش رقمية في عمود {col}")

    d, c = _num(df["debit"]), _num(df["credit"])
    add("negative_amount", df.index[(d < 0) | (c < 0)], "مبلغ سالب في المدين أو الدائن")
    add(
        "debit_and_credit_both_populated",
        df.index[(d.fillna(0) != 0) & (c.fillna(0) != 0)],
        "الحساب فيه مدين ودائن في نفس الصف (راجع إن كان ده مقصود)",
    )
    return issues


def analytical_review(df: pd.DataFrame, threshold_pct: float = 20.0) -> dict:
    """مقارنة الرصيد الحالي (مدين - دائن) بالسنة السابقة، وتعليم التغيرات الكبيرة."""
    if "prior_balance" not in df.columns:
        return {
            "check": "analytical_review",
            "skipped": True,
            "reason": "مفيش عمود prior_balance (رصيد السنة السابقة) في الملف",
        }

    net = _num(df["debit"]).fillna(0) - _num(df["credit"]).fillna(0)
    prior = _num(df["prior_balance"])
    flagged = []
    for i in df.index:
        if pd.isna(prior[i]):
            continue
        cur, prv = float(net[i]), float(prior[i])
        if prv == 0:
            if cur != 0:
                flagged.append((i, cur, prv, None))
            continue
        pct = (cur - prv) / abs(prv) * 100
        if abs(pct) > threshold_pct:
            flagged.append((i, cur, prv, round(pct, 1)))

    return {
        "check": "analytical_review",
        "threshold_pct": threshold_pct,
        "flagged": [
            {
                "excel_row": int(i) + 2,
                "account_code": str(df.at[i, "account_code"]),
                "account_name": str(df.at[i, "account_name"]),
                "current_net": cur,
                "prior": prv,
                "change_pct": pct,  # None = كان صفر وبقى له رصيد
            }
            for i, cur, prv, pct in flagged
        ],
    }
