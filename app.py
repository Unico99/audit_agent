"""واجهة ويب لمساعد المراجعة (Streamlit).

التشغيل:
    streamlit run app.py
    (لو الأمر مش معروف:  python -m streamlit run app.py)
"""
import io
import os
import time
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

from agent import checks, fs, sample_templates, templates  # noqa: E402
from agent.classify import GROUP_LABELS, accounts_table  # noqa: E402
from agent.loop import AuditAgent  # noqa: E402
from agent.mock import MockClient  # noqa: E402
from agent.tools import run_tool  # noqa: E402

st.set_page_config(page_title="مساعد المراجعة", page_icon="🧾", layout="wide")

st.markdown(
    """
    <style>
      .stApp, [data-testid="stChatMessage"], [data-testid="stMarkdownContainer"],
      [data-testid="stAlert"], .stCaption { direction: rtl; text-align: right; }
      [data-testid="stChatInput"] textarea { direction: rtl; text-align: right; }
      code, pre { direction: ltr; text-align: left; }
    </style>
    """,
    unsafe_allow_html=True,
)

MODE_MOCK = "تجريبي (من غير API key)"
MODE_REAL = "كلود الحقيقي"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NONE_OPT = "— غير مخصص —"
GROUP_BY_LABEL = {v: k for k, v in GROUP_LABELS.items()}

# الجداول اللي المراجع بيعدّلها قبل ما يوافق على الكتابة
PENDING_CFG = {
    "write_workpaper": {
        "key": "findings",
        "title": "ورقة العمل",
        "cols": ["area", "finding", "standard_reference", "severity", "recommendation"],
        "config": {
            "area": ("المجال", None),
            "finding": ("الملاحظة", "large"),
            "standard_reference": ("المرجع", None),
            "severity": ("الأهمية", ["high", "medium", "low"]),
            "recommendation": ("التوصية", "large"),
        },
        "text_col": "finding",
    },
    "write_requirements": {
        "key": "items",
        "title": "قائمة متطلبات المراجعة",
        "cols": ["area", "requirement", "evidence", "standard_reference", "priority"],
        "config": {
            "area": ("المجال", None),
            "requirement": ("المتطلب", "large"),
            "evidence": ("المستند المطلوب", "large"),
            "standard_reference": ("المرجع (مقترح)", None),
            "priority": ("الأهمية", ["high", "medium", "low"]),
        },
        "text_col": "requirement",
    },
}

ss = st.session_state
ss.setdefault("chat", [])
ss.setdefault("agent", None)
ss.setdefault("agent_key", None)
ss.setdefault("pending", None)
ss.setdefault("saved_file", None)
ss.setdefault("tools_buf", [])
ss.setdefault("fs_base", None)
ss.setdefault("fs_base_key", None)
ss.setdefault("fs_result", None)


def _dir(env: str, default: str) -> Path:
    p = Path(os.environ.get(env, default))
    p.mkdir(parents=True, exist_ok=True)
    return p


def data_dir() -> Path:
    return _dir("AUDIT_DATA_DIR", "data")


def out_dir() -> Path:
    return _dir("AUDIT_OUTPUT_DIR", "outputs")


def defer_approval(name: str, args: dict) -> bool:
    """الموافقة في الويب بزرار: نحفظ الطلب ونرجّع False لحد ما المراجع يوافق."""
    ss.pending = {"name": name, "args": args}
    return False


def get_agent(mode: str, api_key: str):
    key = (mode, api_key)
    if ss.agent is not None and ss.agent_key == key:
        return ss.agent
    log = lambda line: ss.tools_buf.append(line)  # noqa: E731
    if mode == MODE_MOCK:
        agent = AuditAgent(client=MockClient(), approver=defer_approval, log=log)
    else:
        if not api_key:
            return None
        import anthropic

        agent = AuditAgent(client=anthropic.Anthropic(api_key=api_key), approver=defer_approval, log=log)
    ss.agent, ss.agent_key = agent, key
    return agent


def reset_chat() -> None:
    ss.chat, ss.agent, ss.agent_key = [], None, None
    ss.pending, ss.saved_file = None, None


def tb_status(filename):
    """(متوازن؟, تفاصيل, df) للملف المختار، أو (None, رسالة خطأ, None)."""
    if not filename:
        return None, "", None
    try:
        df = checks.load_trial_balance(data_dir() / filename)
        bal = checks.check_balance(df)
        return bal["balanced"], bal, df
    except Exception as e:
        return None, f"{type(e).__name__}: {e}", None


# ------------------------------------------------------------------ الشريط الجانبي
with st.sidebar:
    st.header("الإعدادات")
    mode = st.radio("وضع التشغيل", [MODE_MOCK, MODE_REAL], key="mode_choice")
    api_key = ""
    if mode == MODE_REAL:
        api_key = st.text_input(
            "ANTHROPIC_API_KEY",
            type="password",
            value=os.environ.get("ANTHROPIC_API_KEY", ""),
            help="بيتحفظ في الجلسة الحالية بس. الأفضل تحطه في ملف .env",
        )
        if not api_key:
            st.warning("محتاج API key من console.anthropic.com")

    st.divider()
    client_name = st.text_input("اسم العميل", key="client_name", placeholder="مثال: شركة النيل")
    period = st.text_input("الفترة", key="period", placeholder="مثال: 2025-12-31")

    st.subheader("الملفات")
    up = st.file_uploader("ارفع ميزان مراجعة", type=["xlsx", "xls", "csv"], key="uploader")
    if up is not None:
        (data_dir() / Path(up.name).name).write_bytes(up.getvalue())
        st.success(f"تم رفع {up.name}")

    files = sorted(f.name for f in data_dir().iterdir() if f.suffix.lower() in (".xlsx", ".xls", ".csv"))
    selected = st.selectbox("ملف الميزان", files, key="selected_file") if files else None

    balanced, bal, tb_df = tb_status(selected)
    if balanced is True:
        st.success(f"✔ الميزان متوازن ({bal['total_debit']:,.2f})")
    elif balanced is False:
        st.error(f"✘ الميزان غير متوازن: الفرق {bal['difference']:,.2f}")
    elif selected:
        st.warning(f"تعذّر قراءة الميزان: {bal}")

    st.caption("فحص الميزان وورقة العمل")
    quick = st.button("فحص الميزان وتجهيز ورقة عمل", key="quick", disabled=not selected, type="primary")
    st.caption("بعد التأكد من التوازن (مقفولة لو غير متوازن)")
    quick_req = st.button("استخراج متطلبات المراجعة", key="quick_req", disabled=balanced is not True)
    quick_fs = st.button("مسودة القوائم المالية", key="quick_fs", disabled=balanced is not True)

    st.divider()
    if st.button("محادثة جديدة", key="reset"):
        reset_chat()
        st.rerun()

st.title("🧾 مساعد المراجعة")
st.caption(
    "كل المخرجات مسودات تحتاج مراجعة واعتماد مراجع مختص. "
    "الحسابات بتتم بكود Python، وكلود بيفسّر النتائج ويجهّز الملاحظات."
)

tab_chat, tab_fs, tab_tpl = st.tabs(["💬 المساعد", "📊 القوائم المالية", "📁 النماذج"])

# ================================================================== المساعد
with tab_chat:
    prompt = st.chat_input("اكتب طلبك، مثلاً: افحص ملف sample_balanced_tb.xlsx")
    if selected:
        if quick:
            prompt = f"افحص ملف {selected} وجهّز ورقة عمل"
        elif quick_req:
            prompt = f"استخرج متطلبات المراجعة لملف {selected}"
        elif quick_fs:
            prompt = f"جهّز مسودة القوائم المالية لملف {selected}"

    if prompt:
        agent = get_agent(mode, api_key)
        if agent is None:
            st.error("اختار الوضع التجريبي، أو حط API key عشان تستخدم كلود الحقيقي.")
        else:
            ss.chat.append({"role": "user", "text": prompt})
            ss.pending, ss.saved_file, ss.tools_buf = None, None, []
            meta_suffix = ""
            if client_name:
                meta_suffix += f" | العميل: {client_name}"
            if period:
                meta_suffix += f" | الفترة: {period}"
            with st.spinner("بشتغل على الطلب..."):
                try:
                    answer = agent.ask(prompt + meta_suffix)
                except Exception as e:
                    answer = f"❌ حصل خطأ: {type(e).__name__}: {e}"
            ss.chat.append({"role": "assistant", "text": answer, "tools": list(ss.tools_buf)})

    for m in ss.chat:
        with st.chat_message(m["role"]):
            st.markdown(m["text"])
            if m.get("tools"):
                with st.expander(f"الأدوات اللي استُخدمت ({len(m['tools'])})"):
                    for line in m["tools"]:
                        st.code(line[:400] + ("…" if len(line) > 400 else ""), language=None)

    if ss.pending:
        name, args = ss.pending["name"], ss.pending["args"]
        cfg = PENDING_CFG[name]
        st.divider()
        st.warning(f"{cfg['title']} جاهزة وبانتظار موافقتك قبل الحفظ. تقدر تعدّل الصفوف الأول.")
        m1, m2 = st.columns(2)
        m1.write(f"**العميل:** {args.get('client') or 'غير محدد'}")
        m2.write(f"**الفترة:** {args.get('period') or 'غير محددة'}")

        df = pd.DataFrame(args.get(cfg["key"], []))
        for c in cfg["cols"]:
            if c not in df.columns:
                df[c] = ""
        colcfg = {}
        for c, (label, extra) in cfg["config"].items():
            if isinstance(extra, list):
                colcfg[c] = st.column_config.SelectboxColumn(label, options=extra)
            else:
                colcfg[c] = st.column_config.TextColumn(label, width=extra) if extra else st.column_config.TextColumn(label)
        edited = st.data_editor(df[cfg["cols"]], key="pending_editor", num_rows="dynamic", use_container_width=True, column_config=colcfg)

        c1, c2, _ = st.columns([1, 1, 3])
        if c1.button("✅ موافق واحفظ", key="approve", type="primary"):
            rows = [r for r in edited.fillna("").to_dict("records") if str(r.get(cfg["text_col"], "")).strip()]
            result, is_error = run_tool(name, {**args, cfg["key"]: rows})
            if is_error:
                st.error(result["error"])
            else:
                ss.saved_file = result["saved"]
                n = result.get("findings_written") or result.get("items_written") or 0
                ss.chat.append({"role": "assistant", "text": f"تم حفظ {cfg['title']} `{result['saved']}` ({n} بند)."})
                if ss.agent is not None:
                    ss.agent.messages.append(
                        {"role": "user", "content": f"[ملاحظة من النظام] المراجع وافق وتم حفظ الملف باسم {result['saved']}."}
                    )
                ss.pending = None
                st.rerun()
        if c2.button("❌ رفض", key="reject"):
            ss.pending = None
            ss.chat.append({"role": "assistant", "text": f"تم رفض {cfg['title']} ولم تُحفظ."})
            st.rerun()

    if ss.saved_file:
        path = out_dir() / ss.saved_file
        if path.exists():
            st.success(f"الملف جاهز: {ss.saved_file}")
            st.download_button("⬇ تحميل الملف (Excel)", data=path.read_bytes(), file_name=path.name, mime=XLSX_MIME, key="download")

# ================================================================== القوائم المالية
with tab_fs:
    st.subheader("مسودة القوائم المالية على نموذج المكتب")
    if not selected:
        st.info("ارفع ميزان مراجعة من الشريط الجانبي.")
    elif balanced is not True:
        st.error("القوائم المالية متاحة بعد التأكد إن الميزان متوازن. صلّح الفرق الأول (راجع تبويب المساعد: فحص الميزان).")
    else:
        tpl_path, tpl_meta = templates.get_active("financial_statements")
        lines = fs.template_lines(tpl_path, tpl_meta["config"])
        by_ref = {l["ref"]: l for l in lines}
        disp = {l["ref"]: f"{l['sheet']} › {l['label']} (#{l['row']})" for l in lines if l["candidate"]}
        ref_by_disp = {v: k for k, v in disp.items()}

        st.caption(
            f"النموذج النشط: **{tpl_meta['name']}**" + (" (افتراضي، ارفع نموذج مكتبك من تبويب النماذج)" if tpl_meta.get("builtin") else "")
        )
        key = fs.mapping_key(client_name, Path(selected).stem)
        saved = fs.load_mapping(key)
        if saved and saved.get("template") != tpl_meta["name"]:
            st.warning("التخصيص المحفوظ كان لنموذج مختلف. راجع التخصيص من جديد واعتمده.")

        acc = accounts_table(tb_df)
        base_key = (selected, key, tpl_meta["name"], bool(saved), ss.get("fs_regen", 0))
        if ss.fs_base_key != base_key:
            use_saved = bool(saved and saved.get("template") == tpl_meta["name"])
            sug = fs.suggest_mapping(acc, lines)
            rows = []
            for _, a in acc.iterrows():
                if use_saved and a["code"] in saved["accounts"]:
                    ref = saved["accounts"][a["code"]].get("line")
                    grp = saved["accounts"][a["code"]].get("group") or a["group"]
                else:
                    ref, grp = sug.get(a["code"], {}).get("line"), a["group"]
                rows.append({"code": a["code"], "name": a["name"], "group": GROUP_LABELS.get(grp, ""), "amount": a["amount"], "line": disp.get(ref, NONE_OPT)})
            ni_ref = saved["accounts"].get(fs.NI_KEY, {}).get("line") if use_saved else fs.suggest_ni_line(lines)
            rows.append({"code": fs.NI_KEY, "name": fs.NI_NAME, "group": "حقوق ملكية", "amount": None, "line": disp.get(ni_ref, NONE_OPT)})
            ss.fs_base, ss.fs_base_key = pd.DataFrame(rows), base_key
            ss.fs_result = None

        st.write(
            "راجع تخصيص كل حساب لبند في النموذج (والمجموعة: أصل/التزام/...)، "
            + ("وده **تخصيص معتمد سابقًا**." if saved and saved.get("approved") else "وده **اقتراح آلي غير معتمد**.")
            + " بعد المراجعة اضغط «اعتماد وحفظ»."
        )
        edited_map = st.data_editor(
            ss.fs_base,
            key="fs_editor",
            use_container_width=True,
            hide_index=True,
            disabled=["code", "name", "amount"],
            column_config={
                "code": st.column_config.TextColumn("الكود"),
                "name": st.column_config.TextColumn("اسم الحساب"),
                "amount": st.column_config.NumberColumn("الرصيد", format="%.2f"),
                "group": st.column_config.SelectboxColumn("المجموعة", options=list(GROUP_BY_LABEL)),
                "line": st.column_config.SelectboxColumn("البند في النموذج", options=[NONE_OPT] + list(disp.values()), width="large"),
            },
        )
        b1, b2, _ = st.columns([1.2, 1.2, 3])
        if b1.button("✅ اعتماد وحفظ التخصيص", key="fs_save", type="primary"):
            accounts_map = {}
            for r in edited_map.to_dict("records"):
                ref = ref_by_disp.get(r["line"])
                entry = {"line": ref}
                if r["code"] != fs.NI_KEY:
                    entry["group"] = GROUP_BY_LABEL.get(r["group"])
                accounts_map[r["code"]] = entry
            fs.save_mapping(
                key,
                {
                    "template": tpl_meta["name"],
                    "approved": True,
                    "approved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "labels": {l["ref"]: l["label"] for l in lines},
                    "accounts": accounts_map,
                },
            )
            ss.fs_base_key = None
            st.success("تم حفظ واعتماد التخصيص للعميل. اضغط الآن «جهّز المسودة».")
            st.rerun()
        if b2.button("🔄 اقتراح آلي من جديد (يمسح التخصيص المحفوظ)", key="fs_regen_btn"):
            m = fs.mapping_key(client_name, Path(selected).stem)
            p = fs.mappings_dir() / f"{m}.json"
            p.unlink(missing_ok=True)
            ss["fs_regen"] = ss.get("fs_regen", 0) + 1
            ss.fs_base_key = None
            st.rerun()

        st.divider()
        if st.button("📄 جهّز مسودة القوائم المالية", key="fs_build", disabled=not (saved and saved.get("approved"))):
            res, is_err = run_tool("build_financial_statements", {"path": selected, "client": client_name, "period": period})
            ss.fs_result = {"error": res["error"]} if is_err else res
        if not (saved and saved.get("approved")):
            st.caption("اعتمد التخصيص أولًا لتفعيل الزرار.")

        r = ss.fs_result
        if r:
            if "error" in r:
                st.error(r["error"])
            elif r.get("ok") is False:
                st.error("تعذّر إعداد القوائم:")
                for e in r["errors"]:
                    st.write(f"- {e}")
            else:
                (st.success if r["all_checks_passed"] else st.warning)(
                    f"تم إعداد المسودة: {r['saved']}  |  صافي الربح {r['net_income']:,.2f}"
                )
                for c in r["checks"]:
                    icon = {"ok": "✔", "fail": "✘", "skipped": "—"}[c["status"]]
                    st.write(f"{icon} {c['name']}: {c['detail']}")
                for ab in r.get("abnormal_balances", []):
                    st.write(f"⚠ رصيد عكسي: {ab['code']} {ab['name']} ({ab['group']}): {ab['amount']:,.2f}")
                for w in r.get("warnings", []):
                    if not w.startswith(("عدم تطابق", "لم يتم التحقق")):
                        st.write(f"⚠ {w}")
                p = out_dir() / r["saved"]
                if p.exists():
                    st.download_button("⬇ تحميل القوائم المالية (Excel)", data=p.read_bytes(), file_name=p.name, mime=XLSX_MIME, key="fs_download")

# ================================================================== النماذج
with tab_tpl:
    st.subheader("نماذج المكتب")
    st.caption("ارفع كل نموذج مرة واحدة (Excel بصيغة xlsx)، ويتخزن ويتستخدم تلقائيًا بعد كده. "
               "ممكن تحط في النموذج {{client}} و{{period}} و{{date}} وهتتعبّى تلقائيًا.")

    for kind, title in templates.KINDS.items():
        st.markdown(f"### {title}")
        names = templates.list_templates(kind)
        active = templates.active_name(kind)
        st.write(f"النموذج النشط: **{active or 'افتراضي (لسه ما اتسجلش نموذج)'}**")

        up_t = st.file_uploader("ارفع النموذج", type=["xlsx", "xlsm"], key=f"tpl_up_{kind}")
        tname = st.text_input("اسم للنموذج (اختياري)", key=f"tpl_name_{kind}")
        if st.button("تسجيل النموذج واعتماده", key=f"register_{kind}", disabled=up_t is None):
            try:
                meta = templates.register(kind, up_t.getvalue(), up_t.name, name=tname or None)
                st.success(f"تم تسجيل «{meta['name']}» واعتماده. راجع الأعمدة المكتشفة تحت.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

        if len(names) > 1:
            pick = st.selectbox("تغيير النموذج النشط", names, index=names.index(active) if active in names else 0, key=f"active_{kind}")
            if st.button("تفعيل", key=f"activate_{kind}"):
                templates.set_active(kind, pick)
                st.rerun()

        if active:
            meta = templates.load_meta(kind, active)
            cfg = meta["config"]
            if kind == "requirements":
                st.write("الأعمدة المكتشفة في قالبك (صحّح الحقل لو الاكتشاف غلط):")
                labels = list(templates.REQ_FIELD_LABELS.values())
                inv = {v: k for k, v in templates.REQ_FIELD_LABELS.items()}
                rows = [
                    {"col": L, "header": cfg["headers"].get(L, ""), "field": templates.REQ_FIELD_LABELS.get(f, "(تجاهل)")}
                    for L, f in cfg["columns"].items()
                ]
                hr = st.number_input("رقم صف العناوين", min_value=1, value=int(cfg["header_row"]), key=f"hr_{kind}")
                ed = st.data_editor(
                    pd.DataFrame(rows), key=f"cfg_{kind}", hide_index=True, disabled=["col", "header"],
                    column_config={"col": "العمود", "header": "العنوان في القالب", "field": st.column_config.SelectboxColumn("الحقل", options=labels)},
                )
                if st.button("حفظ إعدادات القالب", key=f"save_cfg_{kind}"):
                    new = dict(cfg)
                    new["header_row"], new["first_data_row"] = int(hr), int(hr) + 1
                    new["columns"] = {r["col"]: inv.get(r["field"], "skip") for r in ed.to_dict("records")}
                    templates.save_config(kind, active, new)
                    st.success("تم الحفظ")
            else:
                st.write("إعدادات كل ورقة في نموذجك (القوائم غير BS/IS بتتساب زي ما هي):")
                stm_labels = {"BS": "قائمة المركز المالي", "IS": "قائمة الدخل", "other": "أخرى (لا تُملأ)"}
                inv = {v: k for k, v in stm_labels.items()}
                rows = [
                    {"sheet": s, "statement": stm_labels.get(c["statement"], "أخرى (لا تُملأ)"), "label_col": c["label_col"],
                     "current_col": c["current_col"], "prior_col": c.get("prior_col", ""), "first_row": int(c["first_row"])}
                    for s, c in cfg["sheets"].items()
                ]
                ed = st.data_editor(
                    pd.DataFrame(rows), key=f"cfg_{kind}", hide_index=True, disabled=["sheet"],
                    column_config={
                        "sheet": "الورقة",
                        "statement": st.column_config.SelectboxColumn("النوع", options=list(stm_labels.values())),
                        "label_col": "عمود البيان", "current_col": "عمود السنة الحالية",
                        "prior_col": "عمود السنة السابقة", "first_row": st.column_config.NumberColumn("أول صف بيانات", step=1),
                    },
                )
                if st.button("حفظ إعدادات النموذج", key=f"save_cfg_{kind}"):
                    new = {"sheets": {}}
                    for r in ed.to_dict("records"):
                        old = cfg["sheets"][r["sheet"]]
                        new["sheets"][r["sheet"]] = {
                            **old, "statement": inv.get(r["statement"], "other"),
                            "label_col": str(r["label_col"]).strip().upper(), "current_col": str(r["current_col"]).strip().upper(),
                            "prior_col": str(r["prior_col"] or "").strip().upper(), "first_row": int(r["first_row"]),
                        }
                    templates.save_config(kind, active, new)
                    st.success("تم الحفظ. راجع تبويب القوائم المالية.")

        buf = io.BytesIO()
        tmp = out_dir() / f"_sample_{kind}.xlsx"
        (sample_templates.make_requirements_template if kind == "requirements" else sample_templates.make_fs_template)(tmp)
        buf.write(tmp.read_bytes())
        tmp.unlink(missing_ok=True)
        st.download_button("⬇ تنزيل النموذج الافتراضي (للاطلاع على الفكرة)", data=buf.getvalue(), file_name=f"sample_{kind}.xlsx", mime=XLSX_MIME, key=f"dl_default_{kind}")
        st.divider()
