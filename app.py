# ============================================================
# app.py - Streamlit Web Application หลัก
# ตรวจคำผิดและจัดระเบียบรูปแบบรายงานการประชุมวุฒิสภา
# รองรับ Multi-user | 200+ หน้า | 4 กฎการตรวจสอบ
# ============================================================

import time
import logging
import pandas as pd
import streamlit as st

import os
from checker_engine import run_full_check
from config import RULES_CONFIG, MAX_FILE_SIZE_MB, GEMINI_MODELS

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)

# ============================================================
# Page Config (ต้องเป็น st call แรกเสมอ)
# ============================================================
st.set_page_config(
    page_title="ตรวจรายงานการประชุมวุฒิสภา",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "About": "ระบบตรวจคำผิดรายงานการประชุมวุฒิสภา v1.0\nพัฒนาด้วย Streamlit + Gemini AI",
    },
)

# --- โหลด Default Values จาก Streamlit Cloud Secrets หรือ Environment Variables ---
# ลำดับ Priority: Streamlit Secrets > OS Environment > ค่าว่าง
def _get_secret(key: str) -> str:
    try:
        val = st.secrets.get(key, "")
        if val:
            return val
    except Exception:
        pass
    return os.environ.get(key, "")

_DEFAULT_API_KEY    = _get_secret("GEMINI_API_KEY")
_DEFAULT_SHEETS_URL = _get_secret("GOOGLE_SHEETS_URL")

# ============================================================
# Custom CSS
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif;
}

/* Header */
.main-header {
    background: linear-gradient(135deg, #1a237e 0%, #283593 60%, #3949ab 100%);
    color: white;
    padding: 1.8rem 2.5rem;
    border-radius: 14px;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 24px rgba(26,35,126,0.25);
}
.main-header h1 { margin: 0; font-size: 1.75rem; font-weight: 700; }
.main-header p  { margin: 0.35rem 0 0; opacity: 0.85; font-size: 0.9rem; }

/* Stat card */
.stat-card {
    background: white;
    border: 1px solid #e0e0e0;
    border-left: 5px solid #ccc;
    border-radius: 8px;
    padding: 0.9rem 1rem;
    margin-bottom: 0.5rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    text-align: center;
}
.stat-card .num   { font-size: 2rem; font-weight: 700; line-height: 1.1; }
.stat-card .label { font-size: 0.8rem; color: #666; margin-top: 0.2rem; }

/* Issue row */
.issue-row {
    background: #fafafa;
    border: 1px solid #ececec;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    margin-bottom: 0.6rem;
}

/* Badge */
.rule-badge {
    display: inline-block;
    padding: 0.2rem 0.6rem;
    border-radius: 12px;
    font-size: 0.75rem;
    font-weight: 600;
    color: white;
}

/* Words */
.wrong-word   { color: #c0392b; font-weight: 700; font-size: 1.05rem; }
.correct-word { color: #27ae60; font-weight: 700; font-size: 1.05rem; }

/* Snippet */
.snippet-box {
    background: #f5f7ff;
    border: 1px solid #dde;
    border-radius: 6px;
    padding: 0.5rem 0.8rem;
    font-size: 0.83rem;
    line-height: 1.7;
    word-break: break-word;
}

/* Hide footer */
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# Header
# ============================================================
st.markdown("""
<div class="main-header">
    <h1>📋 ระบบตรวจรายงานการประชุมวุฒิสภา</h1>
    <p>ตรวจคำผิด &nbsp;|&nbsp; ศัพท์บัญญัติ &nbsp;|&nbsp; วงเล็บซ้ำ &nbsp;|&nbsp; คำทับศัพท์
    &nbsp;&nbsp;—&nbsp;&nbsp; รองรับไฟล์ Word 200+ หน้า &nbsp;|&nbsp; Multi-user</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# Sidebar
# ============================================================
with st.sidebar:
    st.markdown("## ⚙️ การตั้งค่า")

    # API Key
    with st.expander("🔑 Gemini API Key", expanded=True):
        api_key = st.text_input(
            "API Key",
            type="password",
            placeholder="AIza...",
            help="ขอ Key ได้ที่ https://aistudio.google.com/app/apikey",
            key="api_key",
        )
        if api_key:
            st.success("✓ ตั้งค่า API Key แล้ว")
        else:
            st.warning("ยังไม่มี Key → ข้ามการตรวจ AI")

    # Google Sheets URL
    with st.expander("📊 Google Sheets Database", expanded=True):
        sheets_url = st.text_input(
            "CSV Export URL",
            placeholder="https://docs.google.com/spreadsheets/d/.../export?format=csv",
            help="File → Share → Publish to web → CSV → Copy URL",
            key="sheets_url",
        )
        if sheets_url:
            st.success("✓ ตั้งค่า Sheets แล้ว")
        else:
            st.info("ยังไม่มี URL → ข้ามกฎ Sheets")

    st.divider()

    # เปิด/ปิดกฎ
    st.markdown("### 📏 กฎการตรวจสอบ")
    rule_overrides: dict = {}
    for rk, rc in RULES_CONFIG.items():
        rule_overrides[rk] = st.toggle(
            rc["label"],
            value=rc["enabled"],
            help=rc["description"],
            key=f"toggle_{rk}",
        )

    st.divider()

    # แสดงลำดับโมเดล
    with st.expander("🤖 ลำดับ AI Fallback"):
        for i, m in enumerate(GEMINI_MODELS, 1):
            icon = "🥇" if i == 1 else ("🥈" if i == 2 else "🥉")
            st.markdown(f"{icon} `{m}`")

    # วิธีใช้
    with st.expander("ℹ️ วิธีใช้งาน"):
        st.markdown("""
**ขั้นตอน:**
1. ใส่ Gemini API Key
2. ใส่ Google Sheets CSV URL
3. อัปโหลดไฟล์ .docx
4. กด **เริ่มตรวจสอบ**
5. ดูผลในตาราง → Copy Snippet → Ctrl+F ใน Word

**โครงสร้าง Google Sheets (4 คอลัมน์):**
`incorrect_word` | `correct_word` | `note` | `type`

- type = `vocabulary` สำหรับศัพท์บัญญัติ
- type = `transliteration` สำหรับคำทับศัพท์
        """)

# ============================================================
# File Upload
# ============================================================
col_up, col_info = st.columns([3, 1])
with col_up:
    uploaded_file = st.file_uploader(
        "📂 อัปโหลดไฟล์รายงานการประชุม (.docx)",
        type=["docx"],
        help=f"ขนาดไฟล์สูงสุด {MAX_FILE_SIZE_MB} MB",
    )
with col_info:
    if uploaded_file:
        sz = uploaded_file.size / (1024 * 1024)
        st.metric("ขนาดไฟล์", f"{sz:.1f} MB")

# ============================================================
# Session State
# ============================================================
for key in ["check_results", "paragraphs", "last_filename"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ============================================================
# Run Controls
# ============================================================
if uploaded_file:
    file_size_mb = uploaded_file.size / (1024 * 1024)

    if file_size_mb > MAX_FILE_SIZE_MB:
        st.error(f"❌ ไฟล์ใหญ่เกินกำหนด ({file_size_mb:.1f} MB > {MAX_FILE_SIZE_MB} MB)")
        st.stop()

    # อัปเดต rule config ตาม toggle
    for rk, rv in rule_overrides.items():
        RULES_CONFIG[rk]["enabled"] = rv

    btn_c1, btn_c2, btn_c3 = st.columns([2, 1, 3])
    with btn_c1:
        run_btn = st.button("🚀 เริ่มตรวจสอบ", type="primary", use_container_width=True)
    with btn_c2:
        if st.button("🗑️ ล้างผล", use_container_width=True):
            st.session_state.check_results = None
            st.session_state.paragraphs = None
            st.rerun()

    if run_btn:
        file_bytes = uploaded_file.read()
        st.session_state.last_filename = uploaded_file.name

        progress_bar = st.progress(0, text="เริ่มต้นการตรวจสอบ...")
        status_text  = st.empty()
        start_time   = time.time()

        try:
            paragraphs, issues = run_full_check(
                file_bytes=file_bytes,
                api_key=api_key,
                sheets_url=sheets_url,
                progress_bar=progress_bar,
                status_text=status_text,
            )
            elapsed = time.time() - start_time
            st.session_state.check_results = issues
            st.session_state.paragraphs    = paragraphs
            progress_bar.progress(1.0, text=f"✅ เสร็จสมบูรณ์ใน {elapsed:.1f}s")
            status_text.success(
                f"✅ ตรวจครบ {len(paragraphs)} ย่อหน้า 100% | "
                f"พบปัญหา {len(issues)} รายการ | "
                f"ใช้เวลา {elapsed:.1f} วินาที"
            )
        except Exception as err:
            progress_bar.empty()
            st.error(f"❌ เกิดข้อผิดพลาด: {err}")
            logging.exception("run_full_check error")

# ============================================================
# Results Display
# ============================================================
if st.session_state.check_results is not None:
    issues     = st.session_state.check_results
    paragraphs = st.session_state.paragraphs
    filename   = st.session_state.last_filename

    st.divider()
    st.markdown(f"## 📊 ผลการตรวจสอบ — `{filename}`")

    # Summary statistics
    rule_counts = {}
    for iss in issues:
        lbl = iss["rule_label"]
        rule_counts[lbl] = rule_counts.get(lbl, 0) + 1

    color_map = {rc["label"]: rc["color"] for rc in RULES_CONFIG.values()}
    stat_cols  = st.columns(len(RULES_CONFIG) + 1)

    with stat_cols[0]:
        st.markdown(
            '<div class="stat-card" style="border-left-color:#3949ab">'
            f'<div class="num" style="color:#3949ab">{len(issues)}</div>'
            '<div class="label">ทั้งหมด</div></div>',
            unsafe_allow_html=True,
        )
    for i, (lbl, cnt) in enumerate(rule_counts.items(), 1):
        if i < len(stat_cols):
            col = stat_cols[i]
        else:
            col = stat_cols[-1]
        c = color_map.get(lbl, "#888")
        with col:
            st.markdown(
                f'<div class="stat-card" style="border-left-color:{c}">'
                f'<div class="num" style="color:{c}">{cnt}</div>'
                f'<div class="label">{lbl}</div></div>',
                unsafe_allow_html=True,
            )

    st.caption(
        f"ตรวจสอบครบ **{len(paragraphs)} ย่อหน้า** 100% "
        f"| ไฟล์: `{filename}`"
    )

    # Filters
    st.markdown("### 🔎 กรองผลลัพธ์")
    fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 1])
    with fc1:
        all_labels = sorted(set(i["rule_label"] for i in issues))
        filter_rule = st.multiselect("ประเภทปัญหา", all_labels, default=all_labels)
    with fc2:
        search_word = st.text_input("ค้นหาคำ", placeholder="คำที่ต้องการค้น...")
    with fc3:
        max_para = max((i["para_index"] for i in issues), default=1)
        para_range = st.slider("ช่วงย่อหน้า", 1, max_para, (1, max_para))
    with fc4:
        sort_by = st.selectbox("เรียงตาม", ["ย่อหน้า", "ประเภท"])

    # Apply filters
    filtered = [
        i for i in issues
        if i["rule_label"] in filter_rule
        and para_range[0] <= i["para_index"] <= para_range[1]
        and (not search_word or search_word.lower() in (i["wrong_word"] + i["snippet"]).lower())
    ]
    if sort_by == "ประเภท":
        filtered.sort(key=lambda x: (x["rule_label"], x["para_index"]))

    st.markdown(f"**แสดง {len(filtered)} รายการ** จากทั้งหมด {len(issues)} รายการ")

    # Export
    if filtered:
        df_exp = pd.DataFrame([{
            "ลำดับ":         idx + 1,
            "ย่อหน้าที่":    i["para_index"],
            "~หน้าที่":      i["page_hint"],
            "ประเภท":        i["rule_label"],
            "คำผิด":         i["wrong_word"],
            "คำที่ถูกต้อง":  i["correct_word"],
            "เหตุผล":        i["reason"],
            "Snippet":       i["snippet"].replace("[", "").replace("]", ""),
        } for idx, i in enumerate(filtered)])
        csv_bytes = df_exp.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Export CSV (เปิดใน Excel ได้)",
            data=csv_bytes,
            file_name=f"ผลตรวจ_{filename.replace('.docx','')}.csv",
            mime="text/csv",
        )

    # Results list
    st.markdown("---")
    if not filtered:
        st.success("🎉 ไม่พบปัญหาในช่วงที่กรองเลือก")
    else:
        for idx, issue in enumerate(filtered):
            color  = issue["color"]
            wrong  = issue["wrong_word"]
            correct= issue["correct_word"]
            reason = issue["reason"]
            snippet= issue["snippet"]
            label  = issue["rule_label"]
            para   = issue["para_index"]
            page   = issue["page_hint"]

            # Highlight ใน snippet
            snippet_display = snippet.replace(
                f"[{wrong}]",
                f"<b style='background:#fff3cd;border-bottom:2px solid #e67e22'>{wrong}</b>"
            )
            if f"[{wrong}]" not in snippet:
                snippet_display = snippet.replace(
                    wrong,
                    f"<b style='background:#fff3cd;border-bottom:2px solid #e67e22'>{wrong}</b>",
                    1,
                )

            with st.container():
                c1, c2, c3, c4, c5, c6 = st.columns([0.5, 1.3, 1.5, 1.5, 2.5, 1])

                with c1:
                    st.markdown(f"**#{idx+1}**")

                with c2:
                    st.markdown(
                        f'<span class="rule-badge" style="background:{color}">{label}</span><br>'
                        f'<small>ย่อหน้า <b>{para}</b> | ~หน้า {page}</small>',
                        unsafe_allow_html=True,
                    )

                with c3:
                    st.markdown(
                        f'<div class="wrong-word">❌ {wrong}</div>',
                        unsafe_allow_html=True,
                    )

                with c4:
                    st.markdown(
                        f'<div class="correct-word">✅ {correct}</div>',
                        unsafe_allow_html=True,
                    )

                with c5:
                    st.markdown(
                        f'<div class="snippet-box">{snippet_display}</div>',
                        unsafe_allow_html=True,
                    )

                with c6:
                    if st.button("📋 Copy", key=f"cp_{idx}", use_container_width=True,
                                 help="คัดลอก snippet ไปใช้ Ctrl+F ใน Word"):
                        st.session_state[f"show_cp_{idx}"] = not st.session_state.get(f"show_cp_{idx}", False)

                    if st.session_state.get(f"show_cp_{idx}"):
                        clean_snippet = snippet.replace(f"[{wrong}]", wrong)
                        st.text_area(
                            "Ctrl+A แล้ว Ctrl+C:",
                            value=clean_snippet,
                            height=70,
                            key=f"ta_{idx}",
                            label_visibility="collapsed",
                        )

                if idx < len(filtered) - 1:
                    st.markdown('<hr style="margin:0.4rem 0;border-color:#eee">', unsafe_allow_html=True)

# ============================================================
# Empty State
# ============================================================
elif not uploaded_file:
    st.markdown("""
    <div style="text-align:center;padding:3rem 1rem;color:#aaa">
        <div style="font-size:5rem">📄</div>
        <h3 style="color:#bbb;font-weight:400">อัปโหลดไฟล์ .docx เพื่อเริ่มตรวจสอบ</h3>
        <p>รองรับรายงานการประชุมวุฒิสภา ขนาดสูงสุด 100 MB (200+ หน้า)<br>
        ตรวจ 4 กฎ: คำผิดทั่วไป | ศัพท์บัญญัติ | วงเล็บซ้ำ | คำทับศัพท์</p>
    </div>
    """, unsafe_allow_html=True)
