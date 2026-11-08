"""متطلبات المراجعة: سياق من الميزان + مكتبة بداية + كتابة على قالب المكتب."""
from __future__ import annotations

import json
import os
import re
import time
from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from . import templates
from .classify import accounts_table

CATEGORY_LABELS = {
    "cash": "النقدية والبنوك",
    "receivables": "العملاء والمدينون",
    "inventory": "المخزون",
    "fixed_assets": "الأصول الثابتة",
    "payables": "الموردون والدائنون",
    "loans": "القروض والتمويل",
    "accruals": "المستحقات والمخصصات",
    "equity": "حقوق الملكية",
    "revenue": "الإيرادات",
    "cogs": "تكلفة المبيعات",
    "payroll": "المرتبات والأجور",
    "rent": "الإيجارات",
    "interest": "المصروفات التمويلية",
    "tax": "الضرائب",
}
_LIB_CAT = {
    "accum_dep": "fixed_assets",
    "depreciation_exp": "fixed_assets",
    "equity_capital": "equity",
    "equity_retained": "equity",
}
PRIORITY_AR = {"high": "عالية", "medium": "متوسطة", "low": "منخفضة"}
STANDARD_NOTE = "مقترح من مكتبة البداية، يُراجع مقابل النسخة المرخّصة"


def lib_category(cat: str) -> str:
    return _LIB_CAT.get(cat, cat)


def load_library() -> tuple[list[dict], str]:
    custom = Path(os.environ.get("AUDIT_STANDARDS_DIR", "standards")) / "requirements_library.json"
    path = custom if custom.exists() else Path(__file__).with_name("requirements_library.json")
    data = json.loads(path.read_text(encoding="utf-8"))
    return data["items"], ("مكتبة المكتب المخصّصة" if path == custom else "مكتبة البداية")


def build_context(df: pd.DataFrame, bal: dict) -> dict:
    acc = accounts_table(df)
    cats: dict[str, dict] = {}
    for _, a in acc.iterrows():
        lc = lib_category(a["category"])
        if lc not in CATEGORY_LABELS:
            continue
        c = cats.setdefault(lc, {"category": lc, "label": CATEGORY_LABELS[lc], "accounts": [], "balance": 0.0})
        c["accounts"].append(a["name"])
        if a["amount"] is not None:
            c["balance"] += float(a["amount"])

    lib, lib_name = load_library()
    suggested = []
    for it in lib:
        cat = it["category"]
        if cat != "general" and cat not in cats:
            continue
        info = cats.get(cat, {})
        text = it["requirement"].format(
            accounts="، ".join(info.get("accounts", [])[:3]) or "—",
            balance=f"{abs(info.get('balance', 0)):,.0f}",
        )
        suggested.append(
            {
                "area": it["area"],
                "requirement": text,
                "evidence": it.get("evidence", ""),
                "standard_reference": f"{it['standard']} (مقترح)",
                "priority": it.get("priority", "medium"),
            }
        )

    _, meta = templates.get_active("requirements")
    unclassified = [f"{r['code']} {r['name']}" for _, r in acc.iterrows() if lib_category(r["category"]) not in CATEGORY_LABELS and not str(r["category"]).startswith("other_")]
    return {
        "balance": bal,
        "categories_present": [{k: (round(v, 2) if isinstance(v, float) else v) for k, v in c.items()} for c in cats.values()],
        "suggested_items": suggested,
        "library": lib_name,
        "standards_note": STANDARD_NOTE,
        "template": {"name": meta["name"], "builtin": meta.get("builtin", False)},
        "accounts_not_categorized": unclassified[:20],
    }


def write_requirements(items: list[dict], client: str, period: str, out_dir: Path) -> dict:
    path, meta = templates.get_active("requirements")
    cfg = meta["config"]
    wb = load_workbook(path)
    ws = wb[cfg["sheet"]] if cfg["sheet"] in wb.sheetnames else wb.worksheets[0]

    cols = {column_index_from_string(L): f for L, f in cfg["columns"].items() if f and f != "skip"}
    if not cols:
        raise ValueError("إعدادات أعمدة القالب فاضية. اضبطها من تبويب النماذج")
    first = int(cfg["first_data_row"])
    styles = {c: copy(ws.cell(first, c)._style) for c in cols}

    def row_empty(r: int) -> bool:
        return all(ws.cell(r, c).value in (None, "") for c in range(1, ws.max_column + 1))

    region_end = first
    while region_end <= ws.max_row and not row_empty(region_end):
        region_end += 1
    for r in range(first, region_end):  # مسح صفوف المثال الموجودة في القالب
        for c in cols:
            ws.cell(r, c).value = None
    needed = len(items)
    has_content_below = any(not row_empty(r) for r in range(region_end, ws.max_row + 1))
    if has_content_below and needed > (region_end - first):  # نزّل التوقيعات/الهوامش اللي تحت الجدول
        ws.insert_rows(region_end, needed - (region_end - first))

    for i, it in enumerate(items, start=1):
        r = first + i - 1
        for c, field in cols.items():
            if field == "index":
                val = i
            elif field == "priority":
                val = PRIORITY_AR.get(str(it.get("priority", "")).lower(), it.get("priority", ""))
            elif field == "status":
                val = it.get("status") or "مطلوب"
            else:
                val = it.get(field, "")
            cell = ws.cell(r, c, val)
            cell._style = copy(styles[c])

    templates.fill_placeholders(
        wb, {"client": client or "", "company": client or "", "period": period or "", "date": time.strftime("%Y-%m-%d")}
    )
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"[^\w\-]+", "_", f"متطلبات_المراجعة_{client or 'عميل'}", flags=re.UNICODE).strip("_")
    out = out_dir / f"{stem}.xlsx"
    wb.save(out)
    return {"saved": out.name, "items_written": len(items), "template": meta["name"]}
