"""مسودة القوائم المالية من ميزان المراجعة على نموذج المكتب.

الفكرة:
  1) كل حساب في الميزان يتخصص لبند في نموذجك (اقتراح آلي + اعتماد المراجع).
  2) الكود بيجمّع الأرصدة بإشارتها الطبيعية ويكتبها في خلايا النموذج.
  3) الإجماليات اللي في نموذجك (معادلات Excel) بتفضل زي ما هي، وأنا بحسبها
     بنفسي من المعادلات وأقارنها (الأصول = حقوق الملكية + الالتزامات، وصافي
     الربح = المحسوب من الميزان) عشان أكشف أي تخصيص غلط.
"""
from __future__ import annotations

import difflib
import json
import os
import re
import time
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import column_index_from_string

from . import templates
from .classify import BS_GROUPS, GROUP_LABELS, accounts_table, net_income
from .textutil import norm

NI_KEY = "__NET_INCOME__"
NI_NAME = "صافي ربح/خسارة العام (محسوب)"
_NUM = '#,##0;(#,##0);"-"'

# كلمات بتميّز بنود النموذج لكل فئة حساب (للاقتراح الآلي بس)
CAT_KEYWORDS = {
    "cash": ["نقد", "cash"],
    "receivables": ["عملاء", "مدين", "receivable", "debtors", "customers"],
    "inventory": ["مخزون", "inventor", "stock"],
    "fixed_assets": ["اصول ثابته", "الاصول الثابته", "ممتلكات", "عقارات", "property", "plant", "equipment", "fixed asset"],
    "accum_dep": ["اصول ثابته", "الاصول الثابته", "ممتلكات", "عقارات", "property", "plant", "equipment", "fixed asset"],
    "payables": ["مورد", "دائنون", "payable", "creditors", "suppliers"],
    "accruals": ["مستحق", "accrued", "accrual"],
    "loans": ["قروض", "تمويل", "اقتراض", "loan", "borrowing"],
    "equity_capital": ["راس المال", "capital"],
    "equity_retained": ["محتجز", "ارباح مرحله", "retained"],
    "revenue": ["مبيعات", "ايراد", "revenue", "sales"],
    "cogs": ["تكلفه", "cost of"],
    "payroll": ["مرتب", "اجور", "salar", "wage", "payroll"],
    "rent": ["ايجار", "rent"],
    "depreciation_exp": ["اهلاك", "استهلاك", "depreciation"],
    "interest": ["تمويلي", "فوائد", "finance cost", "interest"],
    "other_expense": ["اخري", "عموميه", "other", "general"],
}
# كلمات في البند تمنع التطابق مع الفئة (مثلاً الإيراد مينفعش يتخصص لبند "تكلفة المبيعات")
CAT_EXCLUDE = {
    "revenue": ["تكلفه", "cost of"],
    "receivables": ["مخصص", "provision"],
    "payables": ["مستحق", "accrued"],
}
NI_LABEL_KEYWORDS = [
    "ارباح العام", "خساير) العام", "نتيجه العام", "صافي ربح العام", "ارباح (خساير) العام",
    "profit for the year", "net profit", "profit/(loss)", "profit (loss)",
]


# ------------------------------------------------------------------ التخزين
def mappings_dir() -> Path:
    p = Path(os.environ.get("AUDIT_MAPPINGS_DIR", "mappings")).resolve()
    p.mkdir(parents=True, exist_ok=True)
    return p


def mapping_key(client: str = "", fallback: str = "") -> str:
    raw = client or fallback or "default"
    return re.sub(r"[^\w\-]+", "_", raw, flags=re.UNICODE).strip("_") or "default"


def load_mapping(key: str) -> dict | None:
    p = mappings_dir() / f"{key}.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None


def save_mapping(key: str, data: dict) -> None:
    (mappings_dir() / f"{key}.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ------------------------------------------------------------------ بنود النموذج
def template_lines(path, config: dict) -> list[dict]:
    wb = load_workbook(path)
    lines: list[dict] = []
    for sheet, sc in config["sheets"].items():
        if sc.get("statement") not in ("BS", "IS") or sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        lc = column_index_from_string(sc["label_col"])
        cc = column_index_from_string(sc["current_col"])
        sheet_lines = []
        for r in range(int(sc["first_row"]), ws.max_row + 1):
            label = ws.cell(r, lc).value
            if not isinstance(label, str) or not label.strip():
                continue
            v = ws.cell(r, cc).value
            is_formula = isinstance(v, str) and v.startswith("=")
            bold = bool(ws.cell(r, lc).font and ws.cell(r, lc).font.b)
            heading = v is None and bold
            sheet_lines.append(
                {
                    "ref": f"{sheet}!{r}",
                    "sheet": sheet,
                    "row": r,
                    "label": label.strip(),
                    "statement": sc["statement"],
                    "formula": is_formula,
                    "candidate": (not is_formula) and (not heading),
                }
            )
        if sheet_lines and not any(l["candidate"] for l in sheet_lines):
            for l in sheet_lines:
                l["candidate"] = not l["formula"]
        lines += sheet_lines
    return lines


# ------------------------------------------------------------------ اقتراح التخصيص
def _score(name: str, category: str, label: str) -> float:
    ln, nn = norm(label), norm(name)
    ratio = difflib.SequenceMatcher(None, nn, ln).ratio()
    kws = CAT_KEYWORDS.get(category, [])
    if any(norm(k) in ln for k in CAT_EXCLUDE.get(category, [])):
        return ratio * 0.5
    if any(norm(k) in ln for k in kws):
        return 1.0 + ratio * 0.1
    return ratio


def suggest_mapping(acc: pd.DataFrame, lines: list[dict]) -> dict:
    """{code: {"line": ref|None, "reason": ...}} اقتراح آلي لازم المراجع يراجعه."""
    out: dict = {}
    for _, a in acc.iterrows():
        if a["group"] is None:
            out[a["code"]] = {"line": None, "reason": "المجموعة غير محددة"}
            continue
        stmt = "BS" if a["group"] in BS_GROUPS else "IS"
        best, best_s = None, 0.0
        for l in lines:
            if l["candidate"] and l["statement"] == stmt:
                s = _score(a["name"], a["category"], l["label"])
                if s > best_s:
                    best, best_s = l, s
        if best is not None and best_s >= 0.6:
            out[a["code"]] = {"line": best["ref"], "reason": "تطابق كلمات" if best_s >= 1 else "تشابه اسم"}
        else:
            out[a["code"]] = {"line": None, "reason": "لا يوجد بند مناسب (اختاره يدويًا)"}
    return out


def suggest_ni_line(lines: list[dict]) -> str | None:
    kws = [norm(k) for k in NI_LABEL_KEYWORDS]
    for l in lines:
        if l["candidate"] and l["statement"] == "BS" and any(k in norm(l["label"]) for k in kws):
            return l["ref"]
    return None


# ------------------------------------------------------------------ مقيّم معادلات Excel (SUM وجمع/طرح)
_TOK = re.compile(
    r"""
    (?P<range>(?:(?:'(?P<qs>[^']+)'|(?P<us>[A-Za-z_؀-ۿ][\w؀-ۿ]*))!)?\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+):\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))
   |(?P<ref>(?:(?:'(?P<qs2>[^']+)'|(?P<us2>[A-Za-z_؀-ۿ][\w؀-ۿ]*))!)?\$?(?P<c>[A-Z]{1,3})\$?(?P<r>\d+))
   |(?P<func>SUM)\(
   |(?P<num>\d+(?:\.\d+)?)
   |(?P<op>[+\-*/(),])
   |(?P<ws>\s+)
    """,
    re.X,
)


class NotSupported(Exception):
    pass


class Evaluator:
    def __init__(self, wb):
        self.wb, self.memo, self.stack = wb, {}, set()

    def cell(self, sheet: str, coord: str) -> float:
        key = (sheet, coord)
        if key in self.memo:
            return self.memo[key]
        if key in self.stack:
            raise NotSupported("مرجع دائري")
        v = self.wb[sheet][coord].value
        if isinstance(v, str) and v.startswith("="):
            self.stack.add(key)
            try:
                val = self.formula(sheet, v[1:])
            finally:
                self.stack.discard(key)
        elif isinstance(v, (int, float)):
            val = float(v)
        else:
            val = 0.0
        self.memo[key] = val
        return val

    def formula(self, sheet: str, f: str) -> float:
        out, pos = [], 0
        while pos < len(f):
            m = _TOK.match(f, pos)
            if not m:
                raise NotSupported(f"صيغة غير مدعومة: {f}")
            pos = m.end()
            if m.group("range"):
                sh = m.group("qs") or m.group("us") or sheet
                out.append(f"_rng({sh!r},{m.group('c1')!r},{int(m.group('r1'))},{m.group('c2')!r},{int(m.group('r2'))})")
            elif m.group("ref"):
                sh = m.group("qs2") or m.group("us2") or sheet
                out.append(f"_ref({sh!r},{m.group('c')!r},{int(m.group('r'))})")
            elif m.group("func"):
                out.append("_sum(")
            elif m.group("num"):
                out.append(m.group("num"))
            elif m.group("op"):
                out.append(m.group("op"))
        expr = "".join(out)

        def _ref(sh, col, row):
            return self.cell(sh, f"{col}{row}")

        def _rng(sh, c1, r1, c2, r2):
            a, b = column_index_from_string(c1), column_index_from_string(c2)
            from openpyxl.utils import get_column_letter as L

            return [self.cell(sh, f"{L(c)}{r}") for c in range(min(a, b), max(a, b) + 1) for r in range(min(r1, r2), max(r1, r2) + 1)]

        def _sum(*args):
            t = 0.0
            for x in args:
                t += sum(x) if isinstance(x, list) else x
            return t

        try:
            return float(eval(expr, {"__builtins__": {}}, {"_ref": _ref, "_rng": _rng, "_sum": _sum}))  # noqa: S307 (تعبير مبني من توكنز مفحوصة فقط)
        except NotSupported:
            raise
        except Exception as e:
            raise NotSupported(f"تعذر حساب المعادلة {f}: {e}") from e


def _find_rows(ws, label_col: int, first_row: int, pred) -> list[int]:
    return [r for r in range(first_row, ws.max_row + 1) if isinstance(ws.cell(r, label_col).value, str) and pred(norm(ws.cell(r, label_col).value))]


_TOTAL_WORDS = ("اجمالي", "مجموع", "total")


def _is_total_assets(n: str) -> bool:
    return (
        any(w in n for w in _TOTAL_WORDS)
        and ("الاصول" in n or "assets" in n)
        and not any(w in n for w in ("متداول", "current", "non-current", "non current", "noncurrent", "liabilities", "equity"))
    )


def _is_total_equity_liabilities(n: str) -> bool:
    return (
        any(w in n for w in _TOTAL_WORDS)
        and ("حقوق الملكيه" in n or "equity" in n)
        and ("الالتزامات" in n or "الخصوم" in n or "liabilities" in n)
    )


def _is_net_profit(n: str) -> bool:
    if "صافي" in n and ("ربح" in n or "خساير" in n or "نتيجه" in n):
        return True
    return any(k in n for k in ("net profit", "net income", "net loss", "profit for the year", "profit/(loss)", "profit (loss)", "loss for the year"))


def template_checks(wb, config: dict, ni_current: float, ni_prior: float | None) -> list[dict]:
    """تحقق من إجماليات النموذج بعد التعبئة (محسوبة من معادلاته)."""
    checks = []
    ev = Evaluator(wb)
    for sheet, sc in config["sheets"].items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        lc = column_index_from_string(sc["label_col"])
        first = int(sc["first_row"])
        cols = [("الحالية", sc["current_col"], ni_current)]
        if sc.get("prior_col") and ni_prior is not None:
            cols.append(("السابقة", sc["prior_col"], ni_prior))
        if sc.get("statement") == "BS":
            a_rows = _find_rows(ws, lc, first, _is_total_assets)
            le_rows = _find_rows(ws, lc, first, _is_total_equity_liabilities)
            if not a_rows or not le_rows:
                checks.append({"name": f"{sheet}: التحقق من توازن القائمة", "status": "skipped", "detail": "لم أجد صفي إجمالي الأصول وإجمالي حقوق الملكية والالتزامات في النموذج"})
                continue
            for lab, col, _ in cols:
                try:
                    a = ev.cell(sheet, f"{col}{a_rows[-1]}")
                    b = ev.cell(sheet, f"{col}{le_rows[-1]}")
                except NotSupported as e:
                    checks.append({"name": f"{sheet}: توازن القائمة ({lab})", "status": "skipped", "detail": str(e)})
                    continue
                checks.append({"name": f"{sheet}: إجمالي الأصول = إجمالي حقوق الملكية والالتزامات ({lab})", "status": "ok" if abs(a - b) <= 0.5 else "fail", "detail": f"{a:,.2f} مقابل {b:,.2f}", "a": a, "b": b})
        elif sc.get("statement") == "IS":
            ni_rows = _find_rows(ws, lc, first, _is_net_profit)
            if not ni_rows:
                checks.append({"name": f"{sheet}: التحقق من صافي الربح", "status": "skipped", "detail": "لم أجد صف صافي الربح في النموذج"})
                continue
            for lab, col, expected in cols:
                try:
                    v = ev.cell(sheet, f"{col}{ni_rows[-1]}")
                except NotSupported as e:
                    checks.append({"name": f"{sheet}: صافي الربح ({lab})", "status": "skipped", "detail": str(e)})
                    continue
                checks.append({"name": f"{sheet}: صافي الربح في النموذج = المحسوب من الميزان ({lab})", "status": "ok" if abs(v - expected) <= 0.5 else "fail", "detail": f"{v:,.2f} مقابل {expected:,.2f}", "a": v, "b": expected})
    return checks


# ------------------------------------------------------------------ البناء
def build(df: pd.DataFrame, mapping: dict, template_path, config: dict, out_dir: Path, client: str = "", period: str = "") -> dict:
    errors: list[str] = []
    warnings: list[str] = []
    overrides = {c: m.get("group") for c, m in mapping.get("accounts", {}).items() if c != NI_KEY and m.get("group")}
    acc = accounts_table(df, overrides)

    dup = acc[acc["code"].duplicated(keep=False) & (acc["code"] != "")]
    if not dup.empty:
        errors.append(f"أكواد حسابات مكررة ({', '.join(sorted(dup['code'].unique()))}): صلّحها في الميزان قبل إعداد القوائم")
    if (acc["code"] == "").any():
        errors.append("في حسابات من غير كود حساب: كمّل الأكواد الأول")

    lines = template_lines(template_path, config)
    by_ref = {l["ref"]: l for l in lines}
    labels = mapping.get("labels", {})
    accts_map = mapping.get("accounts", {})

    def check_line(ref: str, what: str, want_stmt: str) -> bool:
        l = by_ref.get(ref)
        if l is None:
            errors.append(f"{what}: البند {ref} مش موجود في النموذج الحالي (النموذج اتغيّر؟ راجع التخصيص)")
        elif labels.get(ref) and norm(labels[ref]) != norm(l["label"]):
            errors.append(f"{what}: البند في الصف {l['row']} اتغيّر من «{labels[ref]}» إلى «{l['label']}». راجع التخصيص")
        elif not l["candidate"]:
            errors.append(f"{what}: البند «{l['label']}» عنوان أو إجمالي (معادلة) ومينفعش يتكتب فيه رصيد")
        elif l["statement"] != want_stmt:
            errors.append(f"{what}: البند «{l['label']}» في {'قائمة الدخل' if l['statement'] == 'IS' else 'قائمة المركز المالي'} وده مش مكانه")
        else:
            return True
        return False

    totals: dict[str, dict] = {}
    unmapped, no_group = [], []
    has_prior = acc["prior"].notna().any()
    for _, a in acc.iterrows():
        if a["group"] is None:
            no_group.append(f"{a['code']} {a['name']}")
            continue
        m = accts_map.get(a["code"])
        if not m or not m.get("line"):
            unmapped.append(f"{a['code']} {a['name']}")
            continue
        want = "BS" if a["group"] in BS_GROUPS else "IS"
        if check_line(m["line"], f"حساب {a['code']} {a['name']}", want):
            t = totals.setdefault(m["line"], {"current": 0.0, "prior": 0.0})
            t["current"] += float(a["amount"])
            if a["prior"] is not None and not pd.isna(a["prior"]):
                t["prior"] += float(a["prior"])
    if no_group:
        errors.append("حسابات من غير مجموعة (أصل/التزام/...): " + "، ".join(no_group[:10]))
    if unmapped:
        errors.append(f"حسابات غير مخصصة لأي بند ({len(unmapped)}): " + "، ".join(unmapped[:10]) + (" ..." if len(unmapped) > 10 else ""))

    ni_cur = net_income(acc, "amount")
    ni_prior = net_income(acc, "prior") if has_prior else None
    ni_map = accts_map.get(NI_KEY)
    if not ni_map or not ni_map.get("line"):
        errors.append("بند «أرباح (خسائر) العام» في قائمة المركز المالي لازم يتخصص لصافي الربح المحسوب")
    elif check_line(ni_map["line"], "صافي ربح العام", "BS"):
        t = totals.setdefault(ni_map["line"], {"current": 0.0, "prior": 0.0})
        t["current"] += ni_cur
        t["prior"] += ni_prior or 0.0

    if errors:
        return {"ok": False, "errors": errors, "warnings": warnings}

    wb = load_workbook(template_path)
    filled_sheets = set()
    for ref, t in totals.items():
        l = by_ref[ref]
        sc = config["sheets"][l["sheet"]]
        ws = wb[l["sheet"]]
        cell = ws.cell(l["row"], column_index_from_string(sc["current_col"]), round(t["current"], 2))
        if cell.number_format == "General":
            cell.number_format = _NUM
        if sc.get("prior_col") and has_prior:
            pc = ws.cell(l["row"], column_index_from_string(sc["prior_col"]), round(t["prior"], 2))
            if pc.number_format == "General":
                pc.number_format = _NUM
        filled_sheets.add(l["sheet"])
    if not has_prior:
        warnings.append("الميزان لا يحتوي على عمود رصيد السنة السابقة: عمود المقارنة اتساب فاضي")
    for sheet, sc in config["sheets"].items():
        if sheet in wb.sheetnames and sc.get("statement") not in ("BS", "IS"):
            warnings.append(f"ورقة «{sheet}» لم تُملأ (غير مدعومة آليًا): تُعد يدويًا")
    templates.fill_placeholders(wb, {"client": client, "company": client, "period": period, "date": time.strftime("%Y-%m-%d")})

    checks = template_checks(wb, config, ni_cur, ni_prior)
    for c in checks:
        if c["status"] == "fail":
            warnings.append(f"عدم تطابق: {c['name']} ({c['detail']}). غالبًا خطأ في التخصيص")
        elif c["status"] == "skipped":
            warnings.append(f"لم يتم التحقق: {c['name']} ({c['detail']})")

    abnormal = []
    for _, a in acc.iterrows():
        if a["amount"] is None or abs(a["amount"]) < 0.005:
            continue
        if a["amount"] < 0 and a["category"] != "accum_dep":
            abnormal.append({"code": a["code"], "name": a["name"], "group": GROUP_LABELS.get(a["group"], a["group"]), "amount": round(float(a["amount"]), 2)})

    # ورقة ملاحظات المسودة (سجل التخصيص والفحوصات)
    ns = wb.create_sheet("ملاحظات المسودة")
    ns.sheet_view.rightToLeft = True
    ns.append(["مسودة أولية مُجهّزة بمساعدة AI من ميزان المراجعة. لا تُعتمد قبل مراجعة المراجع."])
    ns["A1"].font = Font(bold=True, color="C00000")
    ns.append(["العميل", client])
    ns.append(["الفترة", period])
    ns.append(["صافي الربح المحسوب", round(ni_cur, 2)])
    ns.append([])
    ns.append(["فحوصات النموذج"])
    ns.cell(ns.max_row, 1).font = Font(bold=True)
    for c in checks:
        ns.append([c["name"], {"ok": "✔ مطابق", "fail": "✘ غير مطابق", "skipped": "— لم يتم"}[c["status"]], c["detail"]])
    if abnormal:
        ns.append([])
        ns.append(["أرصدة عكسية (راجع التبويب أو التصنيف)"])
        ns.cell(ns.max_row, 1).font = Font(bold=True)
        for ab in abnormal:
            ns.append([f"{ab['code']} {ab['name']}", ab["group"], ab["amount"]])
    if warnings:
        ns.append([])
        ns.append(["تحذيرات"])
        ns.cell(ns.max_row, 1).font = Font(bold=True)
        for w in warnings:
            ns.append([w])
    ns.append([])
    ns.append(["سجل التخصيص (حساب -> بند)"])
    ns.cell(ns.max_row, 1).font = Font(bold=True)
    ns.append(["كود", "اسم الحساب", "المجموعة", "البند في النموذج", "الرصيد الحالي"])
    for _, a in acc.iterrows():
        ref = accts_map.get(a["code"], {}).get("line")
        ns.append([a["code"], a["name"], GROUP_LABELS.get(a["group"], ""), by_ref[ref]["label"] if ref in by_ref else "", round(float(a["amount"]), 2) if a["amount"] is not None else ""])
    ns.column_dimensions["A"].width = 60
    for col in "BCDE":
        ns.column_dimensions[col].width = 28
    for row in ns.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    wb.calculation.fullCalcOnLoad = True
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"[^\w\-]+", "_", f"القوائم_المالية_{client or 'عميل'}", flags=re.UNICODE).strip("_")
    out = out_dir / f"{stem}.xlsx"
    wb.save(out)

    return {
        "ok": True,
        "saved": out.name,
        "net_income": round(ni_cur, 2),
        "checks": [{k: v for k, v in c.items() if k in ("name", "status", "detail")} for c in checks],
        "all_checks_passed": all(c["status"] == "ok" for c in checks) and bool(checks),
        "abnormal_balances": abnormal,
        "warnings": warnings,
        "sheets_filled": sorted(filled_sheets),
    }
