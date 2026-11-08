"""سجل النماذج: ترفع النموذج مرة واحدة، يتخزن، ويتستخدم تلقائيًا بعد كده.

النماذج بتتخزن في فولدر templates/<النوع>/ ومعاها ملف json فيه إعدادات الأعمدة
المكتشفة (تقدر تعدّلها من الواجهة). لو مفيش نموذج متسجل بيتستخدم نموذج افتراضي.
"""
from __future__ import annotations

import io
import json
import os
import re
import time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from . import sample_templates
from .textutil import norm

KINDS = {
    "requirements": "قالب متطلبات المراجعة",
    "financial_statements": "نموذج القوائم المالية",
}

# حقول متطلبات المراجعة وأسماء العناوين المحتملة في القالب (عربي/إنجليزي)
REQ_FIELDS = {
    "index": ["م", "#", "no", "رقم", "مسلسل", "serial"],
    "standard_reference": ["المرجع", "المعيار", "مرجع المعيار", "reference", "standard"],
    "evidence": ["المستند", "المستندات", "دليل الاثبات", "evidence", "document"],
    "priority": ["الاهميه", "الاولويه", "priority", "importance"],
    "status": ["الحاله", "status"],
    "due": ["تاريخ الاستحقاق", "موعد التسليم", "تاريخ التسليم", "due"],
    "notes": ["ملاحظات", "notes", "comments"],
    "area": ["المجال", "البند", "القسم", "الحساب", "area", "account", "section"],
    "requirement": ["المتطلب", "الاجراء", "البيان", "الوصف", "requirement", "procedure", "description"],
}
REQ_FIELD_LABELS = {
    "index": "م",
    "area": "المجال",
    "requirement": "المتطلب",
    "evidence": "المستند المطلوب",
    "standard_reference": "المرجع",
    "priority": "الأهمية",
    "status": "الحالة",
    "due": "تاريخ الاستحقاق",
    "notes": "ملاحظات",
    "skip": "(تجاهل)",
}
_SHORT_EXACT_ONLY = {"م", "#", "no"}


def templates_dir() -> Path:
    return Path(os.environ.get("AUDIT_TEMPLATES_DIR", "templates")).resolve()


def _kind_dir(kind: str) -> Path:
    if kind not in KINDS:
        raise ValueError(f"نوع نموذج غير معروف: {kind}")
    d = templates_dir() / kind
    d.mkdir(parents=True, exist_ok=True)
    return d


def _safe_name(name: str) -> str:
    return re.sub(r"[^\w\-]+", "_", name or "", flags=re.UNICODE).strip("_") or "template"


# ------------------------------------------------------------------ كشف الهيكل
def match_req_field(text) -> str | None:
    n = norm(text)
    if not n:
        return None
    for field, aliases in REQ_FIELDS.items():  # تطابق تام أولًا
        if any(n == norm(a) for a in aliases):
            return field
    for field, aliases in REQ_FIELDS.items():
        for a in aliases:
            na = norm(a)
            if na in _SHORT_EXACT_ONLY or len(na) < 3:
                continue
            if na in n:
                return field
    return None


def detect_requirements_config(path) -> dict:
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    best_row, best_hits = None, 0
    fallback_row, fallback_cnt = None, 0
    for r in range(1, min(ws.max_row, 40) + 1):
        strs = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        strs = [v for v in strs if isinstance(v, str) and v.strip()]
        hits = sum(1 for v in strs if match_req_field(v))
        if hits > best_hits:
            best_row, best_hits = r, hits
        if len(strs) > fallback_cnt:
            fallback_row, fallback_cnt = r, len(strs)
    if best_hits >= 2:
        hr = best_row
    elif fallback_cnt >= 3:
        hr = fallback_row
    else:
        hr = 1
    cols, used = {}, set()
    for c in range(1, ws.max_column + 1):
        v = ws.cell(hr, c).value
        if not isinstance(v, str) or not v.strip():
            continue
        f = match_req_field(v)
        if f and f not in used:
            cols[get_column_letter(c)] = f
            used.add(f)
        else:
            cols[get_column_letter(c)] = "skip"
    return {
        "sheet": ws.title,
        "header_row": hr,
        "first_data_row": hr + 1,
        "columns": cols,
        "headers": {
            get_column_letter(c): str(ws.cell(hr, c).value).strip()
            for c in range(1, ws.max_column + 1)
            if isinstance(ws.cell(hr, c).value, str) and ws.cell(hr, c).value.strip()
        },
    }


_BS_TITLE = ["مركز", "ميزانيه", "balance", "financial position"]
_IS_TITLE = ["دخل", "ارباح", "income", "profit", "loss", "نتائج"]
_PRIOR_HDR = ["السنه السابقه", "الفتره السابقه", "السابقه", "prior", "previous", "comparative", "مقارن"]
_CUR_HDR = ["السنه الحاليه", "الفتره الحاليه", "الحاليه", "current"]
_YEAR = re.compile(r"^(?:19|20)\d\d$")


def _statement_from_title(title: str) -> str:
    n = norm(title)
    if any(norm(k) in n for k in _BS_TITLE):
        return "BS"
    if any(norm(k) in n for k in _IS_TITLE):
        return "IS"
    return "other"


def _detect_fs_sheet(ws) -> dict:
    max_c = ws.max_column
    cur = prior = None
    header_row = 0
    years = {}
    for r in range(1, min(ws.max_row, 20) + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            s = str(v).strip()
            n = norm(s)
            if any(norm(k) in n for k in _PRIOR_HDR):
                prior, header_row = (prior or c), (header_row or r)
            elif any(norm(k) in n for k in _CUR_HDR):
                cur, header_row = (cur or c), (header_row or r)
            elif _YEAR.match(s):
                years[c] = int(s)
                header_row = header_row or r
    if (cur is None or prior is None) and len(years) >= 2:
        ordered = sorted(years.items(), key=lambda kv: -kv[1])
        cur, prior = ordered[0][0], ordered[1][0]

    start = (header_row or 0) + 1
    str_count, num_count = {}, {}
    for r in range(start, ws.max_row + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and not v.startswith("="):
                str_count[c] = str_count.get(c, 0) + 1
            elif v is not None and (isinstance(v, (int, float)) or str(v).startswith("=")):
                num_count[c] = num_count.get(c, 0) + 1
    label_c = max(str_count, key=str_count.get) if str_count else 1
    if cur is None:
        nums = sorted(c for c in num_count if c != label_c)
        if nums:
            cur = nums[0]
            prior = prior or (nums[1] if len(nums) > 1 else None)
        else:
            cur = label_c + 1
    return {
        "statement": _statement_from_title(ws.title),
        "header_row": header_row,
        "first_row": start,
        "label_col": get_column_letter(label_c),
        "current_col": get_column_letter(cur),
        "prior_col": get_column_letter(prior) if prior else "",
    }


def detect_fs_config(path) -> dict:
    wb = load_workbook(path)
    return {"sheets": {ws.title: _detect_fs_sheet(ws) for ws in wb.worksheets}}


# ------------------------------------------------------------------ السجل
def _meta_path(kind: str, name: str) -> Path:
    return _kind_dir(kind) / f"{name}.json"


def _file_path(kind: str, name: str) -> Path:
    return _kind_dir(kind) / f"{name}.xlsx"


def list_templates(kind: str) -> list[str]:
    return sorted(p.stem for p in _kind_dir(kind).glob("*.json") if not p.name.startswith("_"))


def active_name(kind: str) -> str | None:
    f = _kind_dir(kind) / "_active.txt"
    if f.exists():
        n = f.read_text(encoding="utf-8").strip()
        if n and _meta_path(kind, n).exists():
            return n
    return None


def set_active(kind: str, name: str) -> None:
    if not _meta_path(kind, name).exists():
        raise ValueError(f"النموذج غير موجود: {name}")
    (_kind_dir(kind) / "_active.txt").write_text(name, encoding="utf-8")


def register(kind: str, data: bytes, filename: str, name: str | None = None, make_active: bool = True, builtin: bool = False) -> dict:
    """بيسجّل نموذج Excel جديد (xlsx) ويكتشف أعمدته."""
    if Path(filename).suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("النموذج لازم يكون ملف Excel بصيغة xlsx (احفظه بـ Save As لو كان xls)")
    try:
        load_workbook(io.BytesIO(data))
    except Exception as e:
        raise ValueError(f"الملف مش Excel سليم: {e}") from e

    name = _safe_name(name or Path(filename).stem)
    fp = _file_path(kind, name)
    fp.write_bytes(data)
    config = detect_requirements_config(fp) if kind == "requirements" else detect_fs_config(fp)
    meta = {
        "name": name,
        "kind": kind,
        "original_filename": filename,
        "registered_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "builtin": builtin,
        "config": config,
    }
    _meta_path(kind, name).write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    if make_active:
        set_active(kind, name)
    return meta


def load_meta(kind: str, name: str) -> dict:
    return json.loads(_meta_path(kind, name).read_text(encoding="utf-8"))


def save_config(kind: str, name: str, config: dict) -> None:
    meta = load_meta(kind, name)
    meta["config"] = config
    _meta_path(kind, name).write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def template_path(kind: str, name: str) -> Path:
    return _file_path(kind, name)


def ensure_default(kind: str) -> None:
    if "default" in list_templates(kind):
        return
    tmp = _kind_dir(kind) / "_default_tmp.xlsx"
    if kind == "requirements":
        sample_templates.make_requirements_template(tmp)
    else:
        sample_templates.make_fs_template(tmp)
    try:
        register(kind, tmp.read_bytes(), "default.xlsx", name="default", make_active=False, builtin=True)
    finally:
        tmp.unlink(missing_ok=True)


def get_active(kind: str) -> tuple[Path, dict]:
    """بيرجّع (مسار الملف, meta) للنموذج النشط، وبيعمل الافتراضي لو مفيش حاجة."""
    name = active_name(kind)
    if name is None:
        ensure_default(kind)
        name = "default"
        set_active(kind, name)
    return _file_path(kind, name), load_meta(kind, name)


def fill_placeholders(wb, values: dict) -> None:
    """بيستبدل {{client}} و{{period}} و{{company}} و{{date}} في كل خلايا النموذج."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and "{{" in v:
                    for k, val in values.items():
                        v = v.replace("{{" + k + "}}", str(val))
                    c.value = v
