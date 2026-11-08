"""تعبئة القوائم المالية والإيضاحات ومتطلبات المراجعة على نماذج المكتب.

كل الأرقام بتتحسب بكود Python من ميزان المراجعة (مفيش ذكاء اصطناعي هنا).
الإجماليات بتتكتب كمعادلات Excel عشان تفضل حيّة، والفحوصات المتقاطعة في ورقة «ملاحظات المسودة».
"""
from __future__ import annotations

import json
import re
from copy import copy
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill

from agent.checks import check_balance, check_data_quality, load_trial_balance

CONFIG_PATH = Path(__file__).with_name("config.json")
YELLOW = PatternFill("solid", fgColor="FFFF99")
TEMPLATE_COMPANY = "شركة بيانات للملكية الفكرية"

AR_DIGITS = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def load_config(path: Path | None = None) -> dict:
    return json.loads(Path(path or CONFIG_PATH).read_text(encoding="utf-8"))


# ───────────────────────── قراءة الميزان وبوابة التوازن ─────────────────────────
def _norm_code(v) -> str:
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def group_of(code: str, prefixes: dict) -> tuple[str, str] | None:
    best = None
    for p, val in prefixes.items():
        if code.startswith(p) and (best is None or len(p) > len(best[0])):
            best = (p, val)
    if not best:
        return None
    g, _, tag = best[1].partition(":")
    return g, tag


def balance_hints(df: pd.DataFrame, diff: float) -> list[str]:
    """تلميحات عن سبب الفرق (اقتراحات للمراجعة مش حكم نهائي)."""
    hints = []
    d = pd.to_numeric(df["debit"], errors="coerce").fillna(0)
    c = pd.to_numeric(df["credit"], errors="coerce").fillna(0)
    ad = abs(diff)
    if ad == 0:
        return hints
    for label, ser in (("مدين", d), ("دائن", c)):
        for i in df.index[(ser - ad).abs() <= 0.01]:
            hints.append(f"حساب {df.at[i, 'account_code']} ({df.at[i, 'account_name']}) رصيده {label} يساوي الفرق بالظبط — "
                         f"ممكن يكون مدخل في جهة واحدة أو مفقود قيده المقابل (صف {i + 2}).")
    half = ad / 2
    for label, ser, other in (("مدين", d, "الدائن"), ("دائن", c, "المدين")):
        for i in df.index[(ser - half).abs() <= 0.01]:
            hints.append(f"حساب {df.at[i, 'account_code']} ({df.at[i, 'account_name']}) رصيده {label} = نصف الفرق — "
                         f"ممكن يكون اتسجل في الجهة الغلط (المفروض {other}) (صف {i + 2}).")
    if round(ad * 100) % 9 == 0 and ad >= 9:
        hints.append("الفرق يقبل القسمة على 9: ممكن يكون في خطأ تبديل أرقام (transposition) في رقم ما.")
    if ad < 1:
        hints.append("الفرق أقل من جنيه/ريال: غالبًا تقريب في الكسور.")
    return hints


def analyze_tb(path: str | Path, config: dict | None = None) -> dict:
    """بيقرا الميزان ويرجع: توازن + أسباب + حسابات غير مربوطة. مبيكتبش ملفات."""
    cfg = config or load_config()
    out = {"ok": False, "balanced": False, "reasons": [], "warnings": [], "hints": [], "unmapped": []}
    try:
        df = load_trial_balance(path)
    except Exception as e:  # noqa: BLE001
        out["reasons"].append(f"تعذّر قراءة الملف: {e}")
        return out
    df = df.copy()
    df["account_code"] = df["account_code"].map(lambda v: _norm_code(v) if pd.notna(v) else "")
    bal = check_balance(df)
    issues = check_data_quality(df)
    out.update(total_debit=bal["total_debit"], total_credit=bal["total_credit"],
               difference=bal["difference"], balanced=bal["balanced"], df=df, accounts=len(df))
    if not bal["balanced"]:
        out["reasons"].append(
            f"مجموع المدين {bal['total_debit']:,.2f} ≠ مجموع الدائن {bal['total_credit']:,.2f} — الفرق {bal['difference']:,.2f}")
        out["hints"] = balance_hints(df, bal["difference"])
    blocking = {"missing_account_code", "duplicate_account_code", "non_numeric_debit", "non_numeric_credit",
                "negative_amount"}
    for it in issues:
        txt = f"{it['detail']} — صفوف إكسل: {it['excel_rows']}"
        (out["reasons"] if (it["issue"] in blocking and not bal["balanced"]) or it["issue"] in
         {"non_numeric_debit", "non_numeric_credit"} else out["warnings"]).append(txt)
    if "prior_balance" in df.columns:
        p = pd.to_numeric(df["prior_balance"], errors="coerce").fillna(0)
        if abs(p.sum()) > 0.01:
            out["warnings"].append(f"عمود السنة السابقة مش متوازن (مجموعه {p.sum():,.2f}) — أرقام المقارنة غير موثوقة.")
    else:
        out["warnings"].append("مفيش عمود «السنة السابقة»: هتتملي السنة الحالية بس، وأعمدة المقارنة هتفضل فاضية.")
    for i in df.index:
        if not group_of(df.at[i, "account_code"], cfg["prefixes"]):
            out["unmapped"].append(f"{df.at[i, 'account_code']} {df.at[i, 'account_name']}")
    out["ok"] = bal["balanced"] and not out["reasons"]
    return out


# ───────────────────────── نموذج البيانات المجمّعة ─────────────────────────
class TB:
    def __init__(self, df: pd.DataFrame, cfg: dict):
        d = df.copy()
        d["cur"] = pd.to_numeric(d["debit"], errors="coerce").fillna(0) - pd.to_numeric(d["credit"], errors="coerce").fillna(0)
        self.has_prior = "prior_balance" in d.columns
        d["prv"] = pd.to_numeric(d["prior_balance"], errors="coerce").fillna(0) if self.has_prior else 0.0
        gt = d["account_code"].map(lambda c: group_of(c, cfg["prefixes"]) or ("?", ""))
        d["g"] = [x[0] for x in gt]
        d["tag"] = [x[1] for x in gt]
        self.d = d.sort_values("account_code", kind="stable")

    def rows(self, g: str):
        return self.d[self.d.g == g]

    def s(self, g: str | None = None, tag: str | None = None, col: str = "cur", sign: int = 1) -> float:
        x = self.d
        if g:
            x = x[x.g == g]
        if tag:
            x = x[x.tag == tag]
        return round(sign * float(x[col].sum()), 2)

    PL = ("revenue", "cogs_exp", "mkt_exp", "ga_exp", "fx", "other_income", "zakat_exp")

    def net_income(self, col: str) -> float:
        return round(-float(self.d[self.d.g.isin(self.PL)][col].sum()), 2)


# ───────────────────────── أدوات الإكسل ─────────────────────────
def _year_tokens(wb, year: int):
    def rep(s: str) -> str:
        s = re.sub(r"(?<![\d٠-٩])2023(?![\d٠-٩])", "@@Y@@", s)
        s = re.sub(r"(?<![\d٠-٩])2022(?![\d٠-٩])", "@@P@@", s)
        s = re.sub(r"(?<![\d٠-٩])٢٠٢٣(?![\d٠-٩])", "@@AY@@", s)
        s = re.sub(r"(?<![\d٠-٩])٢٠٢٢(?![\d٠-٩])", "@@AP@@", s)
        return (s.replace("@@Y@@", str(year)).replace("@@P@@", str(year - 1))
                 .replace("@@AY@@", str(year).translate(AR_DIGITS)).replace("@@AP@@", str(year - 1).translate(AR_DIGITS)))
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    c.value = rep(c.value)


def mark(ws, ref: str, value=None):
    c = ws[ref]
    c.value = value
    c.fill = YELLOW


def _shift_merges(ws, after_row: int, delta: int):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row > after_row:
            ws.merged_cells.remove(mr)
            mr.shift(row_shift=delta)
            ws.merged_cells.add(mr)
    dims = {r: (d.height) for r, d in ws.row_dimensions.items() if d.height is not None}
    for r in sorted([r for r in dims if r > after_row], reverse=(delta > 0)):
        ws.row_dimensions[r + delta].height = dims[r]
        ws.row_dimensions[r].height = None


def resize_table(ws, first: int, last: int, n: int) -> int:
    """بيظبط عدد صفوف جدول (first..last) ليبقى n صف. بيرجع آخر صف جديد."""
    n = max(n, 1)
    cap = last - first + 1
    if n > cap:
        delta = n - cap
        ws.insert_rows(last + 1, delta)
        _shift_merges(ws, last, delta)
        for r in range(last + 1, last + 1 + delta):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c)._style = copy(ws.cell(last, c)._style)
        return last + delta
    if n < cap:
        delta = cap - n
        ws.delete_rows(first + n, delta)
        _shift_merges(ws, last, -delta)
        return last - delta
    return last


def _blank(ws, r0, r1, cols):
    for r in range(r0, r1 + 1):
        for c in cols:
            ws.cell(r, c).value = None


def _rep_label_ws(wb, sheet, mapping):
    ws = wb[sheet]
    for ref, txt in mapping.items():
        ws[ref].value = txt


# ───────────────────────── القوائم المالية ─────────────────────────
def fill_financials(tb_df: pd.DataFrame, template: str | Path, out: str | Path, client: str, year: int,
                    cfg: dict | None = None) -> dict:
    cfg = cfg or load_config()
    t = TB(tb_df, cfg)
    wb = openpyxl.load_workbook(template)
    _year_tokens(wb, year)
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and TEMPLATE_COMPANY in c.value:
                    c.value = c.value.replace(TEMPLATE_COMPANY, client)
    P = t.has_prior
    manual: list[str] = []
    checks: list[tuple[str, str, str]] = []  # (اسم الفحص, ✔/✘/⚠, تفاصيل)

    def prv(v):  # قيمة السنة السابقة أو فاضي لو مفيش
        return v if P else None

    # ── قيم مشتقة
    def asset(g, col, **k): return t.s(g, col=col, **k)
    def liab(g, col, **k): return t.s(g, col=col, sign=-1, **k)
    V = {}
    for col, key in (("cur", "c"), ("prv", "p")):
        V[key] = dict(
            cash=asset("cash_hand", col) + asset("cash_bank", col),
            ar_gross=asset("ar", col), allow=liab("ar_allow", col) * -1,  # الرصيد الدائن كقيمة سالبة للأصل
            prepaid=asset("prepaid", col) + asset("staff_adv", col),
            due_from=asset("due_from_rp", col),
            ppe_cost=asset("ppe_cost", col), ppe_dep=asset("ppe_accdep", col),
            ap=liab("ap", col), accrued=liab("accrued", col), due_to=liab("due_to_rp", col),
            zakat=liab("zakat_payable", col), eosb=liab("eosb", col),
            capital=liab("capital", col), reserve=liab("reserve", col), retained=liab("retained", col),
            rev=liab("revenue", col), cogs=asset("cogs_exp", col), mkt=asset("mkt_exp", col), ga=asset("ga_exp", col),
            fx=-asset("fx", col), other=liab("other_income", col), zakat_exp=asset("zakat_exp", col),
            ni=t.net_income(col),
            dep_exp=asset("ga_exp", col, tag="dep_exp"), bad_exp=asset("ga_exp", col, tag="bad_debt_exp"),
            eosb_exp=asset("ga_exp", col, tag="eosb_exp"),
        )
    c, p = V["c"], V["p"]
    for k in (c, p):
        k["ar_net"] = k["ar_gross"] + k["allow"]
        k["ppe_net"] = k["ppe_cost"] + k["ppe_dep"]
        k["re_total"] = k["retained"] + k["ni"]
        k["assets"] = k["cash"] + k["ar_net"] + k["prepaid"] + k["due_from"] + k["ppe_net"]
        k["liab"] = k["ap"] + k["accrued"] + k["due_to"] + k["zakat"] + k["eosb"]
        k["equity"] = k["capital"] + k["reserve"] + k["re_total"]
    eq_gap = round(c["assets"] - c["liab"] - c["equity"], 2)
    checks.append(("الأصول = الالتزامات + حقوق الملكية (السنة الحالية)", "✔" if abs(eq_gap) < 0.01 else "✘",
                   f"الفرق {eq_gap:,.2f}"))
    if P:
        gp = round(p["assets"] - p["liab"] - p["equity"], 2)
        checks.append(("الأصول = الالتزامات + حقوق الملكية (السنة السابقة)", "✔" if abs(gp) < 0.01 else "✘", f"الفرق {gp:,.2f}"))

    # ── قائمة المركز المالي
    n = "قائمة المركز المالي"
    ws = wb[n]
    ws["A34"].value = "الأرباح المبقاة / (الخسائر المتراكمة)"
    for col, d in (("C", c), ("D", p)):
        if col == "D" and not P:
            for r in (9, 10, 11, 12, 16, 21, 22, 23, 24, 28, 32, 33, 34):
                ws[f"D{r}"].value = None
            continue
        for r, k in ((9, "cash"), (10, "ar_net"), (11, "prepaid"), (12, "due_from"), (16, "ppe_net"), (21, "ap"),
                     (22, "accrued"), (23, "due_to"), (24, "zakat"), (28, "eosb"), (32, "capital"), (33, "reserve"),
                     (34, "re_total")):
            ws[f"{col}{r}"].value = d[k]
        ws[f"{col}13"].value = f"=SUM({col}9:{col}12)"
        ws[f"{col}17"].value = f"={col}13+{col}16"
        ws[f"{col}25"].value = f"=SUM({col}21:{col}24)"
        ws[f"{col}29"].value = f"={col}25+{col}28"
        ws[f"{col}35"].value = f"=SUM({col}32:{col}34)"
        ws[f"{col}36"].value = f"={col}29+{col}35"

    # ── قائمة الدخل
    ws = wb["قائمة الدخل الشامل"]
    ws["A13"].value = "صافي الربح / (الخسارة) من العمليات الرئيسية"
    ws["A18"].value = "صافي الربح / (الخسارة) قبل الزكاة"
    ws["A21"].value = "صافي ربح / (خسارة) السنة"
    ws["A24"].value = "مجموع الدخل / (الخسارة) الشاملة للسنة"
    for col, d in (("C", c), ("D", p)):
        if col == "D" and not P:
            for r in (7, 8, 11, 12, 16, 17, 20):
                ws[f"D{r}"].value = None
            continue
        ws[f"{col}7"].value = d["rev"]
        ws[f"{col}8"].value = -d["cogs"]
        ws[f"{col}9"].value = f"={col}7+{col}8"
        ws[f"{col}11"].value = -d["mkt"]
        ws[f"{col}12"].value = -d["ga"]
        ws[f"{col}13"].value = f"={col}9+{col}11+{col}12"
        ws[f"{col}16"].value = d["fx"]
        ws[f"{col}17"].value = d["other"]
        ws[f"{col}18"].value = f"={col}13+{col}16+{col}17"
        ws[f"{col}20"].value = -d["zakat_exp"]
        ws[f"{col}21"].value = f"={col}18+{col}20"
        ws[f"{col}24"].value = f"={col}21"

    # ── قائمة التغير في حقوق الملكية
    ws = wb["قائمة التغيرات في حقوق الملكية"]
    for ref, txt in {"A8": "صافي ربح / (خسارة) السنة", "A9": "الدخل / (الخسارة) الشاملة الآخر للسنة",
                     "A10": "اجمالي الدخل / (الخسارة) الشاملة للسنة", "A13": "صافي ربح / (خسارة) السنة",
                     "A14": "الدخل / (الخسارة) الشاملة الآخر للسنة", "A15": "اجمالي الدخل / (الخسارة) الشاملة للسنة",
                     "D6": "الأرباح المبقاة / (الخسائر المتراكمة) ريال سعودي"}.items():
        ws[ref].value = txt
    if P:
        ws["B7"], ws["C7"], ws["D7"] = p["capital"], p["reserve"], p["retained"]
        ws["E7"] = "=SUM(B7:D7)"
        ws["D8"], ws["E8"] = p["ni"], "=D8"
        ws["B9"] = ws["C9"] = ws["D9"] = ws["E9"] = 0
        ws["B10"], ws["C10"], ws["D10"], ws["E10"] = 0, 0, "=D8+D9", "=SUM(B10:D10)"
        ws["B11"], ws["C11"], ws["D11"], ws["E11"] = "=B7+B10", "=C7+C10", "=D7+D10", "=SUM(B11:D11)"
    else:
        _blank(ws, 7, 11, (2, 3, 4, 5))
    ws["D13"], ws["E13"] = c["ni"], "=D13"
    ws["D14"] = 0
    ws["D15"], ws["E15"] = "=D13+D14", "=D15"
    if P:
        ws["B16"], ws["C16"], ws["D16"] = "=B11+B15", "=C11+C15", "=D11+D15"
    else:
        ws["B16"], ws["C16"], ws["D16"] = c["capital"], c["reserve"], c["retained"] + c["ni"]
        mark(ws, "A12")
    ws["E16"] = "=SUM(B16:D16)"
    if P:
        open_gap = round(p["retained"] + p["ni"] - c["retained"], 2)
        checks.append(("الأرباح المبقاة الافتتاحية = ختامية السنة السابقة", "✔" if abs(open_gap) < 0.01 else "⚠",
                       f"الفرق {open_gap:,.2f} (توزيعات أو تسويات بدون صف في النموذج)" if abs(open_gap) >= 0.01 else "مطابقة"))
        cap_gap = round(p["capital"] - c["capital"], 2)
        res_gap = round(p["reserve"] - c["reserve"], 2)
        if abs(cap_gap) > 0.01 or abs(res_gap) > 0.01:
            checks.append(("حركة رأس المال/الاحتياطي خلال السنة", "⚠",
                           f"تغيّر (رأس المال {-cap_gap:,.2f} / الاحتياطي {-res_gap:,.2f}) ومفيش صف ليه في النموذج"))

    # ── قائمة التدفقات النقدية (طريقة غير مباشرة من فروق أرصدة السنتين)
    ws = wb["قائمة التدفقات النقدية"]
    ws["A7"].value = "صافي الربح / (الخسارة) قبل الزكاة"
    ws["A12"].value = "الربح / (الخسارة) التشغيلي قبل التغيرات في رأس المال العامل"
    if P:
        pre_zakat = c["ni"] + c["zakat_exp"]
        zakat_paid = -(p["zakat"] + c["zakat_exp"] - c["zakat"])
        eosb_paid = -(p["eosb"] + c["eosb_exp"] - c["eosb"])
        capex = -(c["ppe_cost"] - p["ppe_cost"])
        cf = {
            7: pre_zakat, 9: c["bad_exp"], 10: c["eosb_exp"], 11: c["dep_exp"],
            14: -(c["ar_gross"] - p["ar_gross"]), 15: -(c["prepaid"] - p["prepaid"]),
            16: -(c["due_from"] - p["due_from"]), 17: c["ap"] - p["ap"], 18: c["accrued"] - p["accrued"],
            20: zakat_paid, 21: eosb_paid, 25: capex, 29: c["due_to"] - p["due_to"], 33: p["cash"],
        }
        for r, v in cf.items():
            ws[f"B{r}"].value = round(v, 2)
        ws["B12"] = "=B7+SUM(B9:B11)"
        ws["B19"] = "=B12+SUM(B14:B18)"
        ws["B22"] = "=B19+B20+B21"
        ws["B26"] = "=B25"
        ws["B30"] = "=B29"
        ws["B32"] = "=B22+B26+B30"
        ws["B34"] = "=B32+B33"
        calc_close = round(cf[7] + cf[9] + cf[10] + cf[11] + cf[14] + cf[15] + cf[16] + cf[17] + cf[18] + cf[20] + cf[21]
                           + cf[25] + cf[29] + cf[33], 2)
        gap = round(calc_close - c["cash"], 2)
        checks.append(("النقدية آخر السنة في التدفقات = النقدية في المركز المالي", "✔" if abs(gap) < 0.01 else "✘",
                       f"الفرق {gap:,.2f}" + ("" if abs(gap) < 0.01 else " (غالبًا حركة غير ظاهرة: توزيعات/زيادة رأس مال/استبعاد أصول)")))
    else:
        _blank(ws, 7, 34, (2,))
        manual.append("قائمة التدفقات النقدية: تحتاج أرصدة السنة السابقة")
    for r in range(7, 35):
        if ws[f"C{r}"].value is not None:
            mark(ws, f"C{r}")  # بيمسح قيم النموذج القديمة ويظلّل
    if P:
        manual.append("قائمة التدفقات النقدية — عمود السنة السابقة: يحتاج أرصدة السنة التي قبلها (مظلّل بالأصفر)")

    # ── الإيضاحات المبوّبة
    # 5 النقد
    ws = wb["إيضاح 5 - النقد"]
    accs = pd.concat([t.rows("cash_hand"), t.rows("cash_bank")])
    last = resize_table(ws, 8, 9, len(accs))
    for i, (_, a) in enumerate(accs.iterrows()):
        r = 8 + i
        ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = a["account_name"], round(a["cur"], 2), (round(a["prv"], 2) if P else None)
    ws[f"B{last + 1}"], ws[f"C{last + 1}"] = f"=SUM(B8:B{last})", (f"=SUM(C8:C{last})" if P else None)

    # 6 مدينون تجاريون
    ws = wb["إيضاح 6 - مدينون تجاريون"]
    ws["B7"], ws["B9"], ws["B10"] = c["ar_gross"], c["allow"], "=B7+B9"
    ws["C7"], ws["C9"], ws["C10"] = prv(p["ar_gross"]), prv(p["allow"]), ("=C7+C9" if P else None)
    ws["B15"], ws["B14"] = c["bad_exp"], -p["allow"] if P else None
    ws["B16"] = "=B14+B15"
    wo = round((-p["allow"] + c["bad_exp"]) - (-c["allow"]), 2) if P else 0
    if P:
        checks.append(("حركة مخصص الديون: رصيد أول + مكوّن = رصيد آخر (لا شطب)", "✔" if abs(wo) < 0.01 else "⚠",
                       "مطابقة" if abs(wo) < 0.01 else f"فرق {wo:,.2f} (شطب ديون أو تسوية غير ظاهرة في الجدول)"))
        ws["C15"], ws["C16"] = p["bad_exp"], -p["allow"]
        ws["C14"] = "=C16-C15"
        mark(ws, "C14", f"=C16-C15")
        manual.append("إيضاح 6 — رصيد أول مخصص الديون للسنة السابقة مشتق بافتراض عدم الشطب (مظلّل)")
    else:
        _blank(ws, 14, 16, (3,))

    # 7 مصروفات مقدمة
    ws = wb["إيضاح 7 - مصروفات مقدمة"]
    accs = pd.concat([t.rows("prepaid"), t.rows("staff_adv")])
    last = resize_table(ws, 7, 8, len(accs))
    for i, (_, a) in enumerate(accs.iterrows()):
        ws[f"A{7 + i}"], ws[f"B{7 + i}"], ws[f"C{7 + i}"] = a["account_name"], round(a["cur"], 2), (round(a["prv"], 2) if P else None)
    ws[f"B{last + 1}"], ws[f"C{last + 1}"] = f"=SUM(B7:B{last})", (f"=SUM(C7:C{last})" if P else None)

    # 8 أطراف ذات علاقة — جدول المستحق إليهم (الأسفل) الأول عشان إزاحة الصفوف
    ws = wb["إيضاح 8 - أطراف ذات علاقة"]
    unknown_parties = []

    def party_table(first, last_, g, sign):
        rows_ = t.rows(g)
        new_last = resize_table(ws, first, last_, len(rows_))
        for i, (_, a) in enumerate(rows_.iterrows()):
            r = first + i
            info = cfg["parties"].get(a["account_name"])
            if not info:
                unknown_parties.append(a["account_name"])
            ws[f"A{r}"] = a["account_name"]
            ws[f"B{r}"] = info["relation"] if info else None
            ws[f"C{r}"] = info["nature"] if info else None
            if not info:
                mark(ws, f"B{r}"); mark(ws, f"C{r}")
            ws[f"D{r}"] = round(sign * a["cur"], 2)
            ws[f"E{r}"] = round(sign * a["prv"], 2) if P else None
        ws[f"D{new_last + 1}"], ws[f"E{new_last + 1}"] = f"=SUM(D{first}:D{new_last})", (f"=SUM(E{first}:E{new_last})" if P else None)
        return new_last
    party_table(30, 34, "due_to_rp", -1)
    party_table(19, 25, "due_from_rp", 1)
    for r in range(9, 15):
        for col in "ABCD":
            mark(ws, f"{col}{r}")
    manual.append("إيضاح 8 — جدول «أهم المعاملات» يحتاج حركة الأطراف ذات العلاقة خلال السنة (مظلّل)")
    if unknown_parties:
        manual.append("إيضاح 8 — طرف غير معرّف في config.json (طبيعة العلاقة/التعامل): " + "، ".join(unknown_parties))

    # 9 ممتلكات ومعدات
    ws = wb["إيضاح 9 - ممتلكات ومعدات"]
    cost, dep = t.rows("ppe_cost"), t.rows("ppe_accdep")
    if len(cost) == 3 and len(dep) == 3:
        for j, col in enumerate("BCD"):
            a, b = cost.iloc[j], dep.iloc[j]
            ws[f"{col}6"] = a["account_name"]
            op = round(a["prv"], 2) if P else None
            ws[f"{col}8"], ws[f"{col}9"] = op, (round(a["cur"] - a["prv"], 2) if P else None)
            ws[f"{col}10"] = f"={col}8+{col}9" if P else round(a["cur"], 2)
            ws[f"{col}13"], ws[f"{col}14"] = (round(-b["prv"], 2) if P else None), (round(-(b["cur"] - b["prv"]), 2) if P else None)
            ws[f"{col}15"] = f"={col}13+{col}14" if P else round(-b["cur"], 2)
            ws[f"{col}18"] = f"={col}10-{col}15"
            ws[f"{col}19"] = round(a["prv"] + b["prv"], 2) if P else None
        for r in (8, 9, 10, 13, 14, 15, 18, 19):
            ws[f"E{r}"] = f"=SUM(B{r}:D{r})" if (P or r in (10, 15, 18)) else None
        dep_chg = round(-(c["ppe_dep"] - p["ppe_dep"]), 2) if P else None
        if P:
            gap = round(dep_chg - c["dep_exp"], 2)
            checks.append(("استهلاك السنة في إيضاح 9 = مصروف الاستهلاك في الميزان", "✔" if abs(gap) < 0.01 else "⚠",
                           "مطابقة" if abs(gap) < 0.01 else f"فرق {gap:,.2f} (استبعادات أو مصروف استهلاك مصنّف في مكان آخر)"))
    else:
        manual.append(f"إيضاح 9 — النموذج مبني على 3 فئات أصول؛ الميزان فيه {len(cost)} فئة تكلفة و{len(dep)} فئة مجمع. اتساب كما هو للتعبئة اليدوية.")

    # 10 دائنون
    ws = wb["إيضاح 10 - دائنون تجاريون"]
    accs = t.rows("ap")
    last = resize_table(ws, 7, 7, len(accs))
    for i, (_, a) in enumerate(accs.iterrows()):
        ws[f"A{7 + i}"], ws[f"B{7 + i}"], ws[f"C{7 + i}"] = a["account_name"], round(-a["cur"], 2), (round(-a["prv"], 2) if P else None)
    ws[f"B{last + 1}"], ws[f"C{last + 1}"] = f"=SUM(B7:B{last})", (f"=SUM(C7:C{last})" if P else None)

    # 11 مصروفات مستحقة
    ws = wb["إيضاح 11 - مصروفات مستحقة"]
    accs = t.rows("accrued")
    last = resize_table(ws, 7, 12, len(accs))
    for i, (_, a) in enumerate(accs.iterrows()):
        ws[f"A{7 + i}"], ws[f"B{7 + i}"], ws[f"C{7 + i}"] = a["account_name"], round(-a["cur"], 2), (round(-a["prv"], 2) if P else None)
    ws[f"B{last + 1}"], ws[f"C{last + 1}"] = f"=SUM(B7:B{last})", (f"=SUM(C7:C{last})" if P else None)

    # 12 الزكاة
    ws = wb["إيضاح 12 - الزكاة"]
    ws["B16"], ws["B17"], ws["B18"], ws["B19"] = p["zakat"] if P else None, (-(p["zakat"] + c["zakat_exp"] - c["zakat"]) if P else None), c["zakat_exp"], "=SUM(B16:B18)"
    ws["C18"], ws["C19"] = (p["zakat_exp"], p["zakat"]) if P else (None, None)
    ws["B12"], ws["C12"] = c["zakat_exp"], (p["zakat_exp"] if P else None)
    for r in (9, 10, 11):
        mark(ws, f"B{r}"); mark(ws, f"C{r}")
    mark(ws, "C16"); mark(ws, "C17")
    manual.append("إيضاح 12 — الوعاء الزكوي (الحد الأعلى/الأدنى/الوعاء) والموقف الزكوي وحركة السنة السابقة: إدخال يدوي (مظلّل)")

    # 13 منافع الموظفين
    ws = wb["إيضاح 13 - منافع الموظفين"]
    if P:
        ws["B8"], ws["B9"], ws["B10"], ws["B11"] = p["eosb"], c["eosb_exp"], -(p["eosb"] + c["eosb_exp"] - c["eosb"]), "=SUM(B8:B10)"
        gap = round(ws["B8"].value + ws["B9"].value + ws["B10"].value - c["eosb"], 2)
    ws["C9"], ws["C11"] = (p["eosb_exp"], p["eosb"]) if P else (None, None)
    mark(ws, "C8"); mark(ws, "C10")
    manual.append("إيضاح 13 — رصيد أول ومدفوعات السنة السابقة: إدخال يدوي (مظلّل)")
    if P:
        checks.append(("إيضاح 13: المدفوع محسوب كفرق (رصيد أول + مكوّن − رصيد آخر)", "⚠" if ws["B10"].value > 0 else "✔",
                       "المدفوع موجب = رصيد آخر أعلى من المتوقع؛ راجع تسوية المنافع" if ws["B10"].value > 0 else "مقبول"))

    # 14 رأس المال
    ws = wb["إيضاح 14 - رأس المال"]
    accs = t.rows("capital")
    nom = cfg.get("share_nominal_value", 10)
    last = resize_table(ws, 8, 9, len(accs))
    for i, (_, a) in enumerate(accs.iterrows()):
        r = 8 + i
        ws[f"A{r}"], ws[f"B{r}"] = a["account_name"], nom
        ws[f"D{r}"] = round(-a["cur"], 2)
        ws[f"C{r}"] = f"=D{r}/B{r}"
    ws[f"C{last + 1}"], ws[f"D{last + 1}"] = f"=SUM(C8:C{last})", f"=SUM(D8:D{last})"
    ws["A6"].value = re.sub(r"\(\s*200,000\s*\)", f"({c['capital']:,.0f})", ws["A6"].value or "")
    manual.append("إيضاح 14 — القيمة الاسمية للحصة (config.json) وعدد الحصص مشتق = المبلغ ÷ القيمة؛ راجعه مع عقد التأسيس")

    # 15/16/17 مصروفات
    for sheet, g, first, lst in (("إيضاح 17 - مصروفات عمومية", "ga_exp", 7, 25),
                                 ("إيضاح 16 - مصروفات تسويقية", "mkt_exp", 7, 9),
                                 ("إيضاح 15 - تكلفة الإيرادات", "cogs_exp", 7, 8)):
        ws = wb[sheet]
        accs = t.rows(g)
        last = resize_table(ws, first, lst, len(accs))
        for i, (_, a) in enumerate(accs.iterrows()):
            ws[f"A{first + i}"], ws[f"B{first + i}"], ws[f"C{first + i}"] = a["account_name"], round(a["cur"], 2), (round(a["prv"], 2) if P else None)
        ws[f"B{last + 1}"], ws[f"C{last + 1}"] = f"=SUM(B{first}:B{last})", (f"=SUM(C{first}:C{last})" if P else None)

    manual.append("الإيضاحات النصية (1–4): نص النموذج كما هو بعد استبدال السنة واسم الشركة — راجع النشاط والسياسات والرقم الموحد")

    # ── ورقة الملاحظات
    ns = wb.create_sheet("ملاحظات المسودة", 0)
    ns.sheet_view.rightToLeft = True
    ns["A1"] = f"مسودة قوائم مالية — {client} — {year}"
    ns["A1"].font = Font(bold=True, size=14)
    ns["A2"] = "مسودة آلية تحتاج مراجعة واعتماد مراجع مختص. الإجماليات معادلات Excel (فعّل Enable Editing)."
    ns["A4"], ns["B4"], ns["C4"] = "الفحص", "النتيجة", "التفاصيل"
    for i, (a, b, d_) in enumerate(checks):
        ns.cell(5 + i, 1, a); ns.cell(5 + i, 2, b); ns.cell(5 + i, 3, d_)
    r0 = 7 + len(checks)
    ns.cell(r0, 1, "بنود تحتاج إدخال/مراجعة يدوية (مظلّلة بالأصفر)").font = Font(bold=True)
    for i, m in enumerate(manual):
        ns.cell(r0 + 1 + i, 1, "• " + m)
    ns.column_dimensions["A"].width = 70
    ns.column_dimensions["B"].width = 12
    ns.column_dimensions["C"].width = 70
    wb.save(out)
    return {"checks": checks, "manual": manual, "values": V}


# ───────────────────────── متطلبات المراجعة ─────────────────────────
# مراجع مقترحة على مستوى رقم المعيار فقط — راجعها مقابل نسخك المرخّصة
ISA_REFS = [
    ("Journal entries", "ISA 240 / ISA 330"),
    ("Cash and Cash", "ISA 505 / ISA 500 / IAS 7"),
    ("Trade Receivables", "ISA 505 / ISA 540 / IFRS 9 (ECL) / ISA 500"),
    ("Prepayments", "ISA 500 / ISA 330"),
    ("Related party", "ISA 550 / IAS 24"),
    ("Property and equipment", "ISA 500 / ISA 540 / IAS 16"),
    ("Trade Payable", "ISA 505 / ISA 500 (completeness) / IAS 37"),
    ("Accruals", "ISA 500 / ISA 540 / IAS 37"),
    ("End of Service", "ISA 540 / ISA 500 / IAS 19"),
    ("Equity", "ISA 500 / ISA 580 / IAS 1"),
    ("Sales", "ISA 500 / ISA 330 / ISA 240 (revenue) / IFRS 15"),
    ("Cost of Sales", "ISA 500 / ISA 330 / IFRS 15"),
    ("General & Administrative", "ISA 500 / ISA 330"),
    ("Payroll", "ISA 500 / ISA 330 / IAS 19"),
    ("Other Income", "ISA 500 / ISA 330"),
    ("Trial Balance and Financial Statements", "ISA 315 / ISA 500 / IAS 1 / ISA 560 / ISA 580"),
]


def fill_requirements(template: str | Path, out: str | Path, client: str, year: int, add_refs: bool = True) -> dict:
    wb = openpyxl.load_workbook(template)
    # التاريخ في التمبلت 2025
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    s = re.sub(r"(?<![\d٠-٩])2025(?![\d٠-٩])", str(year), c.value)
                    s = re.sub(r"(?<![\d٠-٩])٢٠٢٥(?![\d٠-٩])", str(year).translate(AR_DIGITS), s)
                    c.value = s
    n_rows = 0
    for ws in wb:
        hdr = 3
        if add_refs:
            title = "المرجع المعياري (مقترح)" if ws.title == "العربية" else "Standard reference (suggested)"
            ws.cell(hdr, 5, title)._style = copy(ws.cell(hdr, 4)._style)
            ws.column_dimensions["E"].width = 34
        for r in range(hdr + 1, ws.max_row + 1):
            ref = ws.cell(r, 1).value
            if not ref or not ws.cell(r, 2).value:
                continue
            n_rows += 1
            if add_refs:
                # التصنيف بالإنجليزي في ورقة English؛ للعربي بنقرا من نفس الصف في English
                eng_cat = wb["English"].cell(r, 3).value or ""
                std = next((v for k, v in ISA_REFS if k.lower() in eng_cat.lower()), "ISA 500")
                cell = ws.cell(r, 5, std)
                cell._style = copy(ws.cell(r, 4)._style)
    wb.properties.title = f"Initial audit requirements — {client} — {year}"
    wb.save(out)
    return {"rows": n_rows // 2}
