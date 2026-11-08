"""MockClient: بيحاكي كلود بسيناريو ثابت عشان تجرّب الدورة كاملة من غير API key.

الأدوات بتتنفذ فعليًا على ملفك، والنتايج حقيقية. اللي متحاكى هو "قرارات" كلود بس:
    1) يفحص شكل الملف
    2) يشغّل الفحوصات + يدوّر في المعايير
    3) يكتب ورقة العمل (بعد موافقتك)
    4) يلخّص

ده مش بديل عن كلود الحقيقي: مفيش فهم للسياق ولا تعامل مع ملفات غريبة أو أسئلة حرة.
"""
from __future__ import annotations

import json
import re
from types import SimpleNamespace as NS

DEFAULT_FILE = "sample_trial_balance.xlsx"


def _resp(content, stop_reason):
    return NS(content=content, stop_reason=stop_reason)


def _text(t):
    return NS(type="text", text=t)


def _tool(uid, name, args):
    return NS(type="tool_use", id=uid, name=name, input=args)


def _current_turn(messages):
    """الرسائل من آخر سؤال نصي للمستخدم لحد دلوقتي."""
    last_text = max(
        i for i, m in enumerate(messages) if m["role"] == "user" and isinstance(m["content"], str)
    )
    return messages[last_text]["content"], messages[last_text + 1 :]


def _results_by_tool(turn_messages):
    """بيرجّع {اسم_الأداة: نتيجتها} من الدورة الحالية."""
    names = {}
    for m in turn_messages:
        if m["role"] == "assistant":
            for b in m["content"]:
                if b.type == "tool_use":
                    names[b.id] = b.name
    out = {}
    for m in turn_messages:
        if m["role"] == "user" and isinstance(m["content"], list):
            for r in m["content"]:
                if r.get("type") == "tool_result":
                    out[names.get(r["tool_use_id"], "?")] = json.loads(r["content"])
    return out


def _build_findings(res: dict) -> list[dict]:
    ref = "غير متاح"
    hits = res.get("search_standards", {}).get("results") or []
    if hits:
        ref = f"{hits[0]['source_file']} (جزء {hits[0]['chunk']})"

    findings = []
    tb = res.get("run_trial_balance_checks", {})
    bal = tb.get("balance")
    if bal and not bal["balanced"]:
        findings.append(
            {
                "area": "ميزان المراجعة",
                "finding": (
                    f"فرق {bal['difference']:,.2f} بين إجمالي المدين ({bal['total_debit']:,.2f}) "
                    f"وإجمالي الدائن ({bal['total_credit']:,.2f})."
                ),
                "standard_reference": ref,
                "severity": "high",
                "recommendation": "تتبع مصدر الفرق وتسويته قبل الاعتماد على الميزان.",
            }
        )
    for issue in tb.get("data_quality_issues", []):
        findings.append(
            {
                "area": "جودة البيانات",
                "finding": f"{issue['detail']} (صفوف إكسل: {issue['excel_rows']}).",
                "standard_reference": ref,
                "severity": "medium",
                "recommendation": "مراجعة الصفوف المذكورة مع العميل وتصحيحها أو توثيق السبب.",
            }
        )
    for f in res.get("analytical_review", {}).get("flagged", []):
        pct = "من صفر" if f["change_pct"] is None else f"{f['change_pct']:+.1f}%"
        findings.append(
            {
                "area": "إجراءات تحليلية",
                "finding": (
                    f"حساب {f['account_code']} {f['account_name']}: تغير {pct} "
                    f"(الحالي {f['current_net']:,.0f} مقابل السابق {f['prior']:,.0f})."
                ),
                "standard_reference": ref,
                "severity": "medium",
                "recommendation": "الحصول على تفسير للتغير ومستندات داعمة قبل الاستنتاج.",
            }
        )
    return findings


def _meta(user_text: str):
    cm = re.search(r"العميل:\s*([^|]+?)\s*(?:\||$)", user_text)
    pm = re.search(r"الفترة:\s*([^|]+?)\s*(?:\||$)", user_text)
    return (cm.group(1).strip() if cm else ""), (pm.group(1).strip() if pm else "")


def _unbalanced_msg(res: dict) -> str | None:
    bal = res.get("run_trial_balance_checks", {}).get("balance")
    if bal and not bal["balanced"]:
        return (
            f"الميزان غير متوازن (فرق {bal['difference']:,.2f}). وقّفت هنا: لا يمكن إخراج المتطلبات "
            "أو القوائم المالية قبل تسوية الفرق."
        )
    return None


def _scenario_workpaper(step, filename, turn, uid, client, period):
    if step == 0:
        return _resp(
            [_text("[وضع تجريبي] هبدأ بفحص شكل الملف."), _tool(f"{uid}_a", "inspect_workbook", {"path": filename})],
            "tool_use",
        )
    if step == 1:
        return _resp(
            [
                _tool(f"{uid}_a", "run_trial_balance_checks", {"path": filename}),
                _tool(f"{uid}_b", "analytical_review", {"path": filename, "threshold_pct": 20}),
                _tool(f"{uid}_c", "search_standards", {"query": "ميزان المراجعة توازن المدين والدائن"}),
            ],
            "tool_use",
        )
    res = _results_by_tool(turn)
    if step == 2:
        return _resp(
            [
                _tool(
                    f"{uid}_a",
                    "write_workpaper",
                    {
                        "filename": f"workpaper_{filename.rsplit('.', 1)[0]}",
                        "client": client or "عميل تجريبي (Mock)",
                        "period": period or "غير محدد",
                        "findings": _build_findings(res),
                    },
                )
            ],
            "tool_use",
        )
    wrote = res.get("write_workpaper", {})
    if "saved" in wrote:
        body = f"تم حفظ ورقة العمل: {wrote['saved']} ({wrote['findings_written']} ملاحظة). المسودة محتاجة مراجعتك."
    else:
        body = "ورقة العمل ما اتكتبتش (الأرجح إنك رفضت أو لسه بانتظار موافقتك). النتايج الأساسية فوق."
    return _resp([_text(f"[وضع تجريبي] خلصت الفحوصات. {body}")], "end_turn")


def _scenario_requirements(step, filename, turn, uid, client, period):
    res = _results_by_tool(turn)
    if step == 0:
        return _resp(
            [_text("[وضع تجريبي] هتأكد الأول إن الميزان متوازن."), _tool(f"{uid}_a", "run_trial_balance_checks", {"path": filename})],
            "tool_use",
        )
    stop = _unbalanced_msg(res)
    if stop:
        return _resp([_text(f"[وضع تجريبي] {stop}")], "end_turn")
    if step == 1:
        return _resp([_tool(f"{uid}_a", "get_requirements_context", {"path": filename, "client": client})], "tool_use")
    ctx = res.get("get_requirements_context", {})
    if step == 2 and "suggested_items" in ctx:
        return _resp(
            [
                _tool(
                    f"{uid}_a",
                    "write_requirements",
                    {"path": filename, "client": client, "period": period, "items": ctx["suggested_items"]},
                )
            ],
            "tool_use",
        )
    wrote = res.get("write_requirements", {})
    if "saved" in wrote:
        body = f"تم حفظ قائمة المتطلبات على قالب «{wrote['template']}»: {wrote['saved']} ({wrote['items_written']} متطلب). مراجع المعايير مقترحة وتحتاج مراجعتك."
    elif "error" in ctx:
        body = f"تعذّر: {ctx['error']}"
    else:
        body = "القائمة جاهزة وبانتظار موافقتك (لم تُحفظ بعد)."
    return _resp([_text(f"[وضع تجريبي] {body}")], "end_turn")


def _scenario_fs(step, filename, turn, uid, client, period):
    res = _results_by_tool(turn)
    if step == 0:
        return _resp(
            [_text("[وضع تجريبي] هتأكد الأول إن الميزان متوازن."), _tool(f"{uid}_a", "run_trial_balance_checks", {"path": filename})],
            "tool_use",
        )
    stop = _unbalanced_msg(res)
    if stop:
        return _resp([_text(f"[وضع تجريبي] {stop}")], "end_turn")
    if step == 1:
        return _resp(
            [_tool(f"{uid}_a", "build_financial_statements", {"path": filename, "client": client, "period": period})],
            "tool_use",
        )
    out = res.get("build_financial_statements", {})
    if "error" in out:
        body = out["error"]
    elif out.get("ok") is False:
        body = "تعذّر إعداد القوائم: " + " | ".join(out.get("errors", []))
    else:
        bad = [c for c in out.get("checks", []) if c["status"] == "fail"]
        status = "كل فحوصات الإجماليات مطابقة" if out.get("all_checks_passed") else (f"في {len(bad)} فحص غير مطابق، راجع ورقة الملاحظات" if bad else "بعض الفحوصات لم تتم، راجع ورقة الملاحظات")
        body = f"تم إعداد مسودة القوائم: {out['saved']}. صافي الربح {out['net_income']:,.0f}. {status}."
        if out.get("abnormal_balances"):
            body += f" فيه {len(out['abnormal_balances'])} رصيد عكسي يحتاج مراجعة."
    return _resp([_text(f"[وضع تجريبي] {body}")], "end_turn")


class _Messages:
    def create(self, *, messages, **_ignored):
        user_text, turn = _current_turn(messages)
        step = sum(1 for m in turn if m["role"] == "assistant")
        m = re.search(r"[\w\-\u0600-\u06FF]+\.(?:xlsx|xlsm|xls|csv)", user_text)
        filename = m.group(0) if m else DEFAULT_FILE
        uid = f"mock_{len(messages)}"
        client, period = _meta(user_text)
        if "متطلبات" in user_text:
            fn = _scenario_requirements
        elif "قوائم" in user_text:
            fn = _scenario_fs
        else:
            fn = _scenario_workpaper
        return fn(step, filename, turn, uid, client, period)


class MockClient:
    def __init__(self):
        self.messages = _Messages()
