"""نماذج افتراضية (مثال فقط) بتتستخدم لحد ما ترفع نماذج مكتبك.

استبدلها بنماذجك من تبويب "النماذج" في الواجهة.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

_THIN = Side(style="thin", color="999999")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR_FILL = PatternFill("solid", fgColor="1F4E78")
_SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
_NUM = '#,##0;(#,##0);"-"'

BS_ROWS = [
    ("a_h", "الأصول", "h", None),
    ("nca_h", "الأصول غير المتداولة", "h", None),
    ("fa", "الأصول الثابتة (بالصافي)", "i", None),
    ("nca_t", "إجمالي الأصول غير المتداولة", "t", "={c}{fa}"),
    ("ca_h", "الأصول المتداولة", "h", None),
    ("cash", "النقدية وما في حكمها", "i", None),
    ("ar", "العملاء والمدينون", "i", None),
    ("inv", "المخزون", "i", None),
    ("ca_t", "إجمالي الأصول المتداولة", "t", "=SUM({c}{cash}:{c}{inv})"),
    ("a_t", "إجمالي الأصول", "t", "={c}{nca_t}+{c}{ca_t}"),
    ("gap", None, "b", None),
    ("le_h", "حقوق الملكية والالتزامات", "h", None),
    ("eq_h", "حقوق الملكية", "h", None),
    ("cap", "رأس المال", "i", None),
    ("re", "الأرباح المحتجزة", "i", None),
    ("ni", "أرباح (خسائر) العام", "i", None),
    ("eq_t", "إجمالي حقوق الملكية", "t", "=SUM({c}{cap}:{c}{ni})"),
    ("ncl_h", "الالتزامات غير المتداولة", "h", None),
    ("loans", "قروض وتسهيلات بنكية", "i", None),
    ("cl_h", "الالتزامات المتداولة", "h", None),
    ("ap", "الموردون والدائنون", "i", None),
    ("acc", "مصروفات مستحقة وأرصدة دائنة أخرى", "i", None),
    ("cl_t", "إجمالي الالتزامات المتداولة", "t", "=SUM({c}{ap}:{c}{acc})"),
    ("le_t", "إجمالي حقوق الملكية والالتزامات", "t", "={c}{eq_t}+{c}{loans}+{c}{cl_t}"),
]
IS_ROWS = [
    ("rev_h", "الإيرادات", "h", None),
    ("rev", "صافي المبيعات والإيرادات", "i", None),
    ("cogs", "تكلفة المبيعات", "i", None),
    ("gp", "مجمل الربح", "t", "={c}{rev}-{c}{cogs}"),
    ("exp_h", "المصروفات التشغيلية", "h", None),
    ("pay", "مرتبات وأجور", "i", None),
    ("rent", "إيجارات", "i", None),
    ("dep", "إهلاك", "i", None),
    ("oth", "مصروفات عمومية وإدارية أخرى", "i", None),
    ("op", "صافي الربح التشغيلي", "t", "={c}{gp}-SUM({c}{pay}:{c}{oth})"),
    ("fin", "مصروفات تمويلية", "i", None),
    ("ni", "صافي ربح (خسارة) العام", "t", "={c}{op}-{c}{fin}"),
]
FIRST_ROW = 6


def _fs_sheet(ws, title: str, rows) -> None:
    ws.sheet_view.rightToLeft = True
    ws["B1"] = "{{company}}"
    ws["B2"] = title
    ws["B3"] = "{{period}}"
    for a in ("B1", "B2", "B3"):
        ws[a].font = Font(bold=True, size=13 if a != "B3" else 11)
    for col, text in zip("BCDE", ("البيان", "إيضاح", "السنة الحالية", "السنة السابقة")):
        c = ws[f"{col}5"]
        c.value = text
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = _HDR_FILL
        c.alignment = Alignment(horizontal="center")
    rowmap = {k: FIRST_ROW + i for i, (k, *_r) in enumerate(rows)}
    for key, label, kind, formula in rows:
        r = rowmap[key]
        if kind == "b":
            continue
        ws[f"B{r}"] = label
        for col in "DE":
            cell = ws[f"{col}{r}"]
            cell.number_format = _NUM
            cell.border = _BORDER
            if formula:
                cell.value = formula.format(c=col, **rowmap)
        ws[f"B{r}"].border = _BORDER
        if kind in ("h", "t"):
            for col in "BDE":
                ws[f"{col}{r}"].font = Font(bold=True)
            if kind == "h":
                for col in "BCDE":
                    ws[f"{col}{r}"].fill = _SUB_FILL
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 9
    ws.column_dimensions["D"].width = 17
    ws.column_dimensions["E"].width = 17


def make_fs_template(path: str | Path) -> Path:
    wb = Workbook()
    _fs_sheet(wb.active, "قائمة المركز المالي", BS_ROWS)
    wb.active.title = "قائمة المركز المالي"
    _fs_sheet(wb.create_sheet("قائمة الدخل"), "قائمة الدخل", IS_ROWS)
    wb.save(path)
    return Path(path)


def make_requirements_template(path: str | Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "متطلبات المراجعة"
    ws.sheet_view.rightToLeft = True
    ws["A1"] = "{{client}}"
    ws["A2"] = "قائمة متطلبات المراجعة"
    ws["A3"] = "{{period}}"
    ws["A4"] = "تاريخ الإعداد: {{date}}"
    ws["A1"].font = ws["A2"].font = Font(bold=True, size=13)
    headers = ["م", "المجال", "المتطلب", "المستند المطلوب", "المرجع (المعيار)", "الأهمية", "الحالة", "تاريخ الاستحقاق", "ملاحظات"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(6, i, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = _HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _BORDER
    for i in range(1, len(headers) + 1):  # صف نموذجي منسّق فاضي
        c = ws.cell(7, i)
        c.border = _BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDEFGHI", (5, 20, 55, 38, 20, 11, 12, 15, 25)):
        ws.column_dimensions[col].width = w
    wb.save(path)
    return Path(path)
