"""تعريف الأدوات (schemas اللي كلود بيشوفها) + التنفيذ الفعلي."""
from __future__ import annotations

import os
import re
from datetime import date
from pathlib import Path
from typing import Callable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from . import checks, fs, requirements, templates
from .standards import search_standards

# ---------------------------------------------------------------- schemas
TOOLS = [
    {
        "name": "inspect_workbook",
        "description": (
            "بيعرض أسماء الأوراق (sheets) وأسماء الأعمدة وأول 5 صفوف من ملف Excel/CSV "
            "جوه فولدر البيانات. استخدمه الأول قبل أي فحص عشان تفهم شكل الملف."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "اسم الملف جوه فولدر البيانات"}
            },
            "required": ["path"],
        },
    },
    {
        "name": "run_trial_balance_checks",
        "description": (
            "بيفحص ميزان المراجعة: توازن المدين والدائن + جودة البيانات (أكواد مكررة، "
            "حقول فاضية، قيم مش رقمية، سالب، مدين ودائن في نفس الصف). "
            "كل الحسابات بتتعمل بكود حتمي، فاعتمد على أرقامه."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "tolerance": {
                    "type": "number",
                    "description": "فرق مسموح بيه في التوازن (افتراضي 0.01)",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "analytical_review",
        "description": (
            "مقارنة الرصيد الحالي بالسنة السابقة وتعليم الحسابات اللي تغيرت أكتر من نسبة معينة. "
            "محتاج عمود رصيد السنة السابقة في الملف."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "threshold_pct": {
                    "type": "number",
                    "description": "نسبة التغير اللي تتعلّم (افتراضي 20)",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "search_standards",
        "description": (
            "بحث في نصوص المعايير والإرشادات الموجودة عند المكتب. بيرجّع مقتطفات مع اسم "
            "الملف المصدر. استخدمه قبل أي استشهاد بمعيار."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"query": {"type": "string"}},
            "required": ["query"],
        },
    },
    {
        "name": "get_requirements_context",
        "description": (
            "بيجهّز سياق متطلبات المراجعة من ميزان متوازن: الفئات الموجودة (نقدية، عملاء، مخزون...) "
            "وأرصدتها، وقائمة متطلبات مقترحة من مكتبة البداية (suggested_items). "
            "بيرفض لو الميزان مش متوازن. راجع الأرقام المقترحة وعدّل عليها حسب الحاجة قبل write_requirements. "
            "مراجع المعايير في المكتبة على مستوى رقم المعيار بس ومعلّمة (مقترح)؛ لا تضف مرجعًا من عندك."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}, "client": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "write_requirements",
        "description": (
            "بيكتب قائمة متطلبات المراجعة على قالب المكتب المسجّل (الصيغة المعتمدة). "
            "بيرفض لو الميزان مش متوازن. بتحتاج موافقة المراجع قبل الحفظ. استدعيه مرة واحدة بعد get_requirements_context."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "ملف الميزان (للتحقق من التوازن)"},
                "client": {"type": "string"},
                "period": {"type": "string"},
                "items": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "area": {"type": "string"},
                            "requirement": {"type": "string"},
                            "evidence": {"type": "string", "description": "المستند أو الدليل المطلوب من العميل"},
                            "standard_reference": {"type": "string"},
                            "priority": {"type": "string", "enum": ["high", "medium", "low"]},
                        },
                        "required": ["area", "requirement", "priority"],
                    },
                },
            },
            "required": ["path", "items"],
        },
    },
    {
        "name": "build_financial_statements",
        "description": (
            "بيعمل مسودة القوائم المالية على نموذج المكتب المعتمد من ميزان متوازن، باستخدام تخصيص الحسابات "
            "اللي المراجع اعتمده (تبويب القوائم المالية). بيرفض لو الميزان مش متوازن أو مفيش تخصيص معتمد. "
            "بيرجّع نتيجة فحص الإجماليات: لو status=fail لا تقل إن القوائم سليمة؛ بلّغ المراجع بالفرق."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "client": {"type": "string"},
                "period": {"type": "string"},
            },
            "required": ["path"],
        },
    },
    {
        "name": "write_workpaper",
        "description": (
            "بيكتب ورقة عمل Excel فيها الملاحظات. بتحتاج موافقة المراجع قبل التنفيذ. "
            "استدعيه مرة واحدة بعد ما تخلّص الفحوصات."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "اسم ملف الإخراج بدون مسار"},
                "client": {"type": "string"},
                "period": {"type": "string", "description": "الفترة، مثلاً 2025-12-31"},
                "findings": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "area": {"type": "string"},
                            "finding": {"type": "string"},
                            "standard_reference": {
                                "type": "string",
                                "description": "مرجع من نتائج search_standards فقط، أو 'غير متاح'",
                            },
                            "severity": {"type": "string", "enum": ["high", "medium", "low"]},
                            "recommendation": {"type": "string"},
                        },
                        "required": ["area", "finding", "severity", "recommendation"],
                    },
                },
            },
            "required": ["filename", "findings"],
        },
    },
]


# ---------------------------------------------------------------- helpers
def _dir(env: str, default: str) -> Path:
    return Path(os.environ.get(env, default)).resolve()


def _safe_path(name: str, base: Path) -> Path:
    """بيمنع الخروج برّه الفولدر المسموح (path traversal)."""
    p = (base / name).resolve()
    if base != p and base not in p.parents:
        raise ValueError("المسار خارج الفولدر المسموح")
    return p


# ---------------------------------------------------------------- implementations
def _inspect_workbook(path: str) -> dict:
    import pandas as pd

    p = _safe_path(path, _dir("AUDIT_DATA_DIR", "data"))
    if p.suffix.lower() == ".csv":
        sheets = {"csv": pd.read_csv(p)}
    else:
        sheets = pd.read_excel(p, sheet_name=None)
    return {
        name: {
            "columns": [str(c) for c in df.columns],
            "rows": len(df),
            "head": df.head(5).astype(str).to_dict(orient="records"),
        }
        for name, df in sheets.items()
    }


def _run_tb_checks(path: str, tolerance: float = 0.01) -> dict:
    df = checks.load_trial_balance(_safe_path(path, _dir("AUDIT_DATA_DIR", "data")))
    return {
        "balance": checks.check_balance(df, tolerance),
        "data_quality_issues": checks.check_data_quality(df),
    }


def _analytical(path: str, threshold_pct: float = 20.0) -> dict:
    df = checks.load_trial_balance(_safe_path(path, _dir("AUDIT_DATA_DIR", "data")))
    return checks.analytical_review(df, threshold_pct)


def _search(query: str) -> dict:
    hits = search_standards(query)
    if not hits:
        return {"results": [], "note": "مفيش نتائج في المعايير المتاحة"}
    return {"results": hits}


def _load_balanced(path: str):
    """بوابة التوازن: أي مخرجات (متطلبات/قوائم) مرفوضة لو الميزان مش متوازن. الفحص بالكود مش بالـ prompt."""
    df = checks.load_trial_balance(_safe_path(path, _dir("AUDIT_DATA_DIR", "data")))
    bal = checks.check_balance(df)
    if not bal["balanced"]:
        raise ValueError(
            f"الميزان غير متوازن (فرق {bal['difference']:,.2f}: مدين {bal['total_debit']:,.2f} مقابل دائن {bal['total_credit']:,.2f}). "
            "لا يمكن المتابعة قبل تسوية الفرق."
        )
    return df, bal


def _get_req_context(path: str, client: str = "") -> dict:
    df, bal = _load_balanced(path)
    return requirements.build_context(df, bal)


def _write_requirements(path: str, items: list, client: str = "", period: str = "") -> dict:
    _load_balanced(path)
    if not items:
        raise ValueError("قائمة المتطلبات فاضية")
    return requirements.write_requirements(items, client, period, _dir("AUDIT_OUTPUT_DIR", "outputs"))


def _build_fs(path: str, client: str = "", period: str = "") -> dict:
    df, _ = _load_balanced(path)
    key = fs.mapping_key(client, Path(path).stem)
    mapping = fs.load_mapping(key)
    if not mapping or not mapping.get("approved"):
        raise ValueError(
            "لا يوجد تخصيص معتمد للحسابات على بنود النموذج لهذا العميل. "
            "اطلب من المراجع يفتح تبويب «القوائم المالية»، يراجع التخصيص ويعتمده، وبعدها أعد المحاولة."
        )
    tpl_path, meta = templates.get_active("financial_statements")
    res = fs.build(df, mapping, tpl_path, meta["config"], _dir("AUDIT_OUTPUT_DIR", "outputs"), client, period)
    res["template"] = meta["name"]
    return res


def _write_workpaper(filename: str, findings: list, client: str = "", period: str = "") -> dict:
    out_dir = _dir("AUDIT_OUTPUT_DIR", "outputs")
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"[^\w\-]+", "_", Path(filename).stem, flags=re.UNICODE) or "workpaper"
    out = _safe_path(f"{stem}.xlsx", out_dir)

    wb = Workbook()
    ws = wb.active
    ws.title = "الملاحظات"
    ws.sheet_view.rightToLeft = True

    ws.append(["العميل", client])
    ws.append(["الفترة", period])
    ws.append(["تاريخ التجهيز", date.today().isoformat()])
    ws.append(["الحالة", "مسودة مُجهّزة بمساعدة AI — لا تُعتمد قبل مراجعة المراجع"])
    ws.append([])

    headers = ["م", "المجال", "الملاحظة", "المرجع", "الأهمية", "التوصية", "حالة المراجعة"]
    ws.append(headers)
    hdr_row = ws.max_row
    for c in ws[hdr_row]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")

    for i, f in enumerate(findings, start=1):
        ws.append(
            [
                i,
                f.get("area", ""),
                f.get("finding", ""),
                f.get("standard_reference", "غير متاح"),
                f.get("severity", ""),
                f.get("recommendation", ""),
                "بانتظار المراجعة",
            ]
        )
    for col, width in zip("ABCDEFG", (5, 22, 55, 28, 10, 45, 18)):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=hdr_row + 1):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out)
    return {"saved": out.name, "findings_written": len(findings)}


# ---------------------------------------------------------------- dispatcher
DISPATCH: dict[str, Callable[..., dict]] = {
    "inspect_workbook": _inspect_workbook,
    "run_trial_balance_checks": _run_tb_checks,
    "analytical_review": _analytical,
    "search_standards": _search,
    "get_requirements_context": _get_req_context,
    "write_requirements": _write_requirements,
    "build_financial_statements": _build_fs,
    "write_workpaper": _write_workpaper,
}

# الأدوات اللي بتكتب ملفات: لازم موافقة بشرية.
# (build_financial_statements مش هنا: البوابة البشرية فيها هي اعتماد المراجع للتخصيص قبل البناء)
NEEDS_APPROVAL = {"write_workpaper", "write_requirements"}


def run_tool(name: str, args: dict, approver: Callable[[str, dict], bool] | None = None):
    """بينفّذ أداة وبيرجّع (النتيجة, هل_حصل_خطأ)."""
    fn = DISPATCH.get(name)
    if fn is None:
        return {"error": f"أداة غير معروفة: {name}"}, True
    if name in NEEDS_APPROVAL and approver is not None and not approver(name, args):
        return {
            "error": "لم تُنفَّذ العملية: المراجع لم يوافق بعد (رفض أو بانتظار الموافقة). "
            "لا تعيد المحاولة؛ اذكر للمراجع إن الورقة جاهزة وبانتظار موافقته."
        }, True
    try:
        return fn(**args), False
    except Exception as e:  # نرجّع الخطأ لكلود عشان يصحّح أو يسأل
        return {"error": f"{type(e).__name__}: {e}"}, True
