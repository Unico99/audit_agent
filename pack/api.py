"""خدمة HTTP صغيرة بتخلي n8n يستدعي كود Python.

التشغيل:  uvicorn pack.api:app --host 0.0.0.0 --port 8000
أمان: لو API_TOKEN متضبط في البيئة، كل طلب لازم يبعت هيدر  X-API-Token  بنفس القيمة.
"""
from __future__ import annotations

import io
import os
import re
import tempfile
import zipfile
from pathlib import Path

from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import JSONResponse, Response

from pack.engine import analyze_tb, fill_financials, fill_requirements, load_config

HERE = Path(__file__).resolve().parent
TPL = Path(os.getenv("PACK_TEMPLATES_DIR", HERE / "templates"))
app = FastAPI(title="Audit pack service")


def _auth(token: str | None):
    want = os.getenv("API_TOKEN")
    if want and token != want:
        raise HTTPException(401, "invalid token")


def _year(period: str) -> int:
    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", period or "")
    if not m:
        raise HTTPException(422, "اكتب الفترة وفيها السنة بأربع أرقام (مثال: 2025)")
    return int(m.group(1))


async def _save(upload: UploadFile, d: Path) -> Path:
    p = d / "tb.xlsx"
    p.write_bytes(await upload.read())
    return p


def _summary(r: dict) -> dict:
    return {k: r.get(k) for k in ("balanced", "total_debit", "total_credit", "difference", "accounts",
                                 "reasons", "hints", "warnings", "unmapped")}


def _message(r: dict) -> str:
    if r["balanced"] and not r["unmapped"] and r["ok"]:
        s = ("✔ الميزان متوازن\nإجمالي المدين: {total_debit:,.2f}\nإجمالي الدائن: {total_credit:,.2f}\n"
             "عدد الحسابات: {accounts}").format(**r)
        if r["warnings"]:
            s += "\n\nتنبيهات:\n- " + "\n- ".join(r["warnings"])
        return s
    lines = ["✘ تم إيقاف العملية — الميزان مش جاهز"]
    if r.get("total_debit") is not None:
        lines.append(f"إجمالي المدين: {r['total_debit']:,.2f}\nإجمالي الدائن: {r['total_credit']:,.2f}\nالفرق: {r['difference']:,.2f}")
    if r["reasons"]:
        lines.append("الأسباب:\n- " + "\n- ".join(r["reasons"]))
    if r["hints"]:
        lines.append("أسباب محتملة للفرق (راجعها):\n- " + "\n- ".join(r["hints"]))
    if r["unmapped"]:
        lines.append("حسابات غير مربوطة بالنماذج (أضفها في config.json):\n- " + "\n- ".join(r["unmapped"]))
    if r["warnings"]:
        lines.append("تنبيهات:\n- " + "\n- ".join(r["warnings"]))
    return "\n\n".join(lines)


@app.get("/health")
def health():
    return {"ok": True, "templates": sorted(p.name for p in TPL.glob("*.xlsx"))}


@app.post("/check")
async def check(file: UploadFile = File(...), x_api_token: str | None = Header(None)):
    _auth(x_api_token)
    with tempfile.TemporaryDirectory() as d:
        r = analyze_tb(await _save(file, Path(d)), load_config())
    out = _summary(r)
    out["ok"] = bool(r["ok"] and not r["unmapped"])
    out["message"] = _message(r)
    return JSONResponse(out)


@app.post("/build")
async def build(file: UploadFile = File(...), client: str = Form(...), period: str = Form(...),
                x_api_token: str | None = Header(None)):
    _auth(x_api_token)
    year = _year(period)
    cfg = load_config()
    with tempfile.TemporaryDirectory() as d:
        d = Path(d)
        r = analyze_tb(await _save(file, d), cfg)
        if not (r["ok"] and not r["unmapped"]):
            return JSONResponse({**_summary(r), "ok": False, "message": _message(r)}, status_code=409)
        fs = d / "fs.xlsx"
        rq = d / "rq.xlsx"
        res = fill_financials(r["df"], TPL / "financial_statements.xlsx", fs, client, year, cfg)
        fill_requirements(TPL / "audit_requirements.xlsx", rq, client, year)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(fs, f"القوائم_المالية_{year}.xlsx")
            z.write(rq, f"متطلبات_المراجعة_{year}.xlsx")
    return Response(buf.getvalue(), media_type="application/zip",
                    headers={"Content-Disposition": f'attachment; filename="audit_pack_{year}.zip"',
                             "X-Checks-Failed": str(sum(1 for c in res["checks"] if c[1] == "✘"))})
