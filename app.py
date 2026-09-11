# ============================================================
# app.py - Streamlit Web Application หลัก
# ตรวจคำผิดและจัดระเบียบรูปแบบรายงานการประชุมวุฒิสภา
# 100% Deterministic Rule & Database Engine — ไม่ใช้ AI
# รองรับ Multi-user | 200+ หน้า | 7 กฎการตรวจสอบ
# ============================================================

import time
import logging
import os
import html
import difflib
import unicodedata
import math
import struct
import base64
import io
import pandas as pd
import streamlit as st

from checker_engine import run_full_check
import config
from config import RULES_CONFIG, MAX_FILE_SIZE_MB


@st.cache_resource
def get_soft_chime_b64() -> str:
    """สร้างไฟล์เสียง WAV สังเคราะห์ 2 โน้ต (D5 -> A5) นุ่มนวล ละมุน สบายหู สำหรับแจ้งเตือนเมื่อตรวจเสร็จ"""
    sample_rate = 44100
    duration = 1.1
    num_samples = int(sample_rate * duration)
    buf = io.BytesIO()

    # RIFF / WAV Header
    num_channels = 1
    bits_per_sample = 16
    byte_rate = sample_rate * num_channels * (bits_per_sample // 8)
    block_align = num_channels * (bits_per_sample // 8)
    data_size = num_samples * block_align

    buf.write(b'RIFF')
    buf.write(struct.pack('<I', 36 + data_size))
    buf.write(b'WAVEfmt ')
    buf.write(struct.pack('<I', 16))
    buf.write(struct.pack('<H', 1))
    buf.write(struct.pack('<H', num_channels))
    buf.write(struct.pack('<I', sample_rate))
    buf.write(struct.pack('<I', byte_rate))
    buf.write(struct.pack('<H', block_align))
    buf.write(struct.pack('<H', bits_per_sample))
    buf.write(b'data')
    buf.write(struct.pack('<I', data_size))

    for i in range(num_samples):
        t = i / sample_rate
        val = 0.0
        # Note 1: 587.33 Hz (D5) - soft decay
        if t >= 0:
            env1 = math.exp(-5.5 * t)
            val += 0.22 * math.sin(2.0 * math.pi * 587.33 * t) * env1
            val += 0.06 * math.sin(2.0 * math.pi * 1174.66 * t) * env1
        # Note 2: 880.0 Hz (A5) - starts at 0.18s
        if t >= 0.18:
            t2 = t - 0.18
            env2 = math.exp(-4.5 * t2)
            val += 0.26 * math.sin(2.0 * math.pi * 880.0 * t2) * env2
            val += 0.08 * math.sin(2.0 * math.pi * 1760.0 * t2) * env2
        val = max(-1.0, min(1.0, val))
        buf.write(struct.pack('<h', int(val * 32767)))

    return base64.b64encode(buf.getvalue()).decode('ascii')


def is_combining_mark(ch: str) -> bool:
    return unicodedata.category(ch) in ('Mn', 'Mc')

def get_thai_clusters(text: str) -> list:
    clusters = []
    curr = ''
    for ch in text:
        if not curr:
            curr = ch
        elif is_combining_mark(ch):
            curr += ch
        else:
            clusters.append(curr)
            curr = ch
    if curr:
        clusters.append(curr)
    return clusters

def render_diff_html(wrong: str, correct: str) -> tuple:
    """สร้าง HTML ไฮไลต์เปรียบเทียบจุดต่างระหว่างคำผิดและคำถูกระดับพยางค์/อักขระ
    เพื่อความชัดเจน ไม่ให้สระบน/ล่าง วรรณยุกต์ลอย หรือเกิดวงกลมประ"""
    if not wrong or not correct:
        return html.escape(str(wrong)), html.escape(str(correct))
    w_clusters = get_thai_clusters(str(wrong))
    c_clusters = get_thai_clusters(str(correct))
    sm = difflib.SequenceMatcher(None, w_clusters, c_clusters)
    
    if sm.ratio() < 0.2:
        w_html = f'<span style="background:#ffcdd2;color:#b71c1c;padding:1px 5px;border-radius:4px;font-weight:700;">{html.escape(str(wrong))}</span>'
        c_html = f'<span style="background:#c8e6c9;color:#1b5e20;padding:1px 5px;border-radius:4px;font-weight:700;">{html.escape(str(correct))}</span>'
        return w_html, c_html

    w_out, c_out = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        w_part = html.escape(''.join(w_clusters[i1:i2]))
        c_part = html.escape(''.join(c_clusters[j1:j2]))
        if tag == 'equal':
            w_out.append(w_part)
            c_out.append(c_part)
        else:
            if w_part:
                w_out.append(f'<span style="background:#ffcdd2;color:#b71c1c;padding:1px 4px;border-radius:3px;font-weight:800;border-bottom:2px solid #d32f2f;box-shadow:0 1px 2px rgba(0,0,0,0.08);">{w_part}</span>')
            if c_part:
                c_out.append(f'<span style="background:#c8e6c9;color:#1b5e20;padding:1px 4px;border-radius:3px;font-weight:800;border-bottom:2px solid #2e7d32;box-shadow:0 1px 2px rgba(0,0,0,0.08);">{c_part}</span>')
    return ''.join(w_out), ''.join(c_out)

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
    initial_sidebar_state="collapsed",
    menu_items={
        "About": "ระบบตรวจคำผิดรายงานการประชุมวุฒิสภา v2.0\nพัฒนาด้วย Streamlit + Rule & Database Engine (100% Deterministic)\nไม่ใช้ AI — ผลลัพธ์แม่นยำ รวดเร็ว ไม่เกิด Hallucination",
    },
)

# ============================================================
# Custom CSS
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');

html, body, [class*="css"], .stMarkdown, .stText, h1, h2, h3, h4, p, label, input, textarea {
    font-family: 'Sarabun', 'Leelawadee UI', 'Tahoma', 'Thonburi', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}

/* คืนค่าฟอนต์ไอคอนให้ Streamlit ป้องกันตัวอักษร upload หรือไอคอนอื่นทับซ้อน */
[data-testid*="stIcon"],
[class*="material-symbols"],
[class*="material-icons"],
.material-symbols-rounded,
.material-icons,
button span,
[data-testid="stFileUploadDropzone"] span {
    font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
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
.wrong-box {
    background: #fff5f5;
    border: 1px solid #fed7d7;
    border-radius: 6px;
    padding: 6px 10px;
}
.correct-box {
    background: #f0fff4;
    border: 1px solid #c6f6d5;
    border-radius: 6px;
    padding: 6px 10px;
}
.word-label {
    font-size: 0.72rem;
    color: #718096;
    margin-bottom: 3px;
    font-weight: 600;
}
.wrong-word   { color: #c53030; font-weight: 700; font-size: 1.05rem; word-break: break-word; }
.correct-word { color: #22543d; font-weight: 700; font-size: 1.05rem; word-break: break-word; }

/* Page Header card */
.page-header-box {
    background: #e8eaf6;
    border: 1px solid #c5cae9;
    border-left: 4px solid #1a237e;
    border-radius: 6px;
    padding: 4px 8px;
    margin-bottom: 6px;
}
.page-header-title {
    font-size: 0.68rem;
    color: #5c6bc0;
    font-weight: 600;
    text-transform: uppercase;
}
.page-header-code {
    font-size: 0.95rem;
    color: #1a237e;
    font-weight: 700;
}

/* Snippet */
.snippet-box {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 0.5rem 0.8rem;
    font-size: 0.83rem;
    line-height: 1.7;
    word-break: break-word;
}

/* ขยายพื้นที่ Drag and Drop Zone ของ File Uploader ให้ใหญ่เต็มพื้นที่ */
[data-testid="stFileUploadDropzone"] {
    min-height: 250px !important;
    border: 2.5px dashed #3949ab !important;
    background-color: #f8faff !important;
    border-radius: 14px !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    align-items: center !important;
    padding: 2.5rem 1.5rem !important;
    margin-top: 0.5rem !important;
    margin-bottom: 1.5rem !important;
    cursor: pointer !important;
    transition: all 0.25s ease-in-out !important;
}
[data-testid="stFileUploadDropzone"]:hover {
    border-color: #1a237e !important;
    background-color: #eef2ff !important;
    box-shadow: 0 4px 16px rgba(57, 73, 171, 0.12) !important;
}
[data-testid="stFileUploadDropzone"] button {
    margin-top: 10px !important;
    padding: 0.5rem 1.5rem !important;
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
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
        <h1 style="margin: 0;">📋 ระบบตรวจรายงานการประชุมวุฒิสภา</h1>
        <span style="font-size: 0.82rem; background: rgba(255, 255, 255, 0.18); padding: 4px 12px; border-radius: 20px; font-weight: 500; letter-spacing: 0.3px;">
            🕒 อัปเดตล่าสุด: 11 ก.ย. 2569 | 14:15 น. (v2.5 Professional)
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# Config (ไม่ต้องใช้ AI API Key อีกต่อไป — 100% Rule & Database Engine)
# ============================================================

# ============================================================
# File Upload
# ============================================================
uploaded_file = st.file_uploader(
    "📂 อัปโหลดไฟล์รายงานการประชุม (.docx)",
    type=["docx"],
    help=f"ขนาดไฟล์สูงสุด {MAX_FILE_SIZE_MB} MB",
)

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

    # กฎทุกข้อเปิดใช้งานอัตโนมัติ (ไม่ต้อง toggle)

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

            # เล่นเสียงแจ้งเตือนแบบละมุน ไม่ดังมาก แต่ได้ยินชัดเจน
            chime_b64 = get_soft_chime_b64()
            st.markdown(
                f"""
                <audio autoplay style="display:none;">
                    <source src="data:audio/wav;base64,{chime_b64}" type="audio/wav">
                </audio>
                <script>
                try {{
                    var snd = new Audio("data:audio/wav;base64,{chime_b64}");
                    snd.volume = 0.35;
                    snd.play();
                }} catch(e) {{}}
                </script>
                """,
                unsafe_allow_html=True
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

    if not issues:
        st.success("✅ ตรวจสอบเอกสารแล้ว ไม่พบข้อผิดพลาดด้านตัวสะกดและข้อเท็จจริงครับ")
        filtered = []
    else:
        # Filters
        st.markdown("### 🔎 กรองผลลัพธ์")
        fc1, fc2, fc3 = st.columns([2, 2, 2])
        with fc1:
            all_labels = sorted(set(i["rule_label"] for i in issues))
            filter_rule = st.multiselect("ประเภทปัญหา", all_labels, default=all_labels)
        with fc2:
            page_codes_ordered = []
            for i in issues:
                pc = i.get("page_code") or f"~หน้า {i.get('page_hint', 1)}"
                if pc not in page_codes_ordered:
                    page_codes_ordered.append(pc)
            filter_page = st.selectbox("📄 กรองตามหัวแผ่นกระดาษ", ["ทุกหน้า (ทั้งหมด)"] + page_codes_ordered)
        with fc3:
            search_word = st.text_input("🔍 ค้นหาคำ / ข้อความ", placeholder="คำผิด, คำถูก, หรือข้อความ...")

        fc4, fc5 = st.columns([3, 1])
        with fc4:
            max_para = max((i["para_index"] for i in issues), default=1)
            if max_para > 1:
                para_range = st.slider("ช่วงย่อหน้า", 1, max_para, (1, max_para))
            else:
                para_range = (1, 1)
        with fc5:
            sort_by = st.selectbox("เรียงตาม", ["ย่อหน้า", "หัวแผ่น / หน้าที่", "ประเภท"])

        # Apply filters
        filtered = [
            i for i in issues
            if i["rule_label"] in filter_rule
            and (filter_page == "ทุกหน้า (ทั้งหมด)" or (i.get("page_code") or f"~หน้า {i.get('page_hint', 1)}") == filter_page)
            and para_range[0] <= i["para_index"] <= para_range[1]
            and (not search_word or search_word.lower() in (
                i["wrong_word"] + i["correct_word"] + i["snippet"] + i.get("page_code", "")
            ).lower())
        ]
        if sort_by == "ประเภท":
            filtered.sort(key=lambda x: (x["rule_label"], x["para_index"]))
        elif sort_by == "หัวแผ่น / หน้าที่":
            filtered.sort(key=lambda x: (x.get("page_code", ""), x["para_index"]))
        else:
            filtered.sort(key=lambda x: x["para_index"])

        st.markdown(f"**แสดง {len(filtered)} รายการ** จากทั้งหมด {len(issues)} รายการ")

    # Export
    if filtered:
        df_exp = pd.DataFrame([{
            "ลำดับ":                 idx + 1,
            "หัวแผ่นกระดาษ / หน้าที่": i.get("page_code", f"~หน้า {i.get('page_hint', 1)}"),
            "หัวแผ่นฉบับเต็ม":         i.get("full_header", ""),
            "ย่อหน้าที่":            i["para_index"],
            "ประเภท":                i["rule_label"],
            "คำผิด":                 i["wrong_word"],
            "คำที่ถูกต้อง":          i["correct_word"],
            "เหตุผล":                i["reason"],
            "ข้อความแวดล้อม (Snippet)": i["snippet"].replace("[", "").replace("]", ""),
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
        st.info("ℹ️ ไม่พบปัญหาตามเงื่อนไขที่เลือกกรอง")
    else:
        # จัดกลุ่มหน้าที่พบปัญหา
        unique_pages = []
        for i in filtered:
            pc = i.get("page_code") or f"หน้า {i.get('page_hint', 1)}"
            if pc not in unique_pages:
                unique_pages.append(pc)
        
        if len(unique_pages) > 1:
            pages_str = ", ".join(unique_pages[:-1]) + f" และ {unique_pages[-1]}"
        elif len(unique_pages) == 1:
            pages_str = unique_pages[0]
        else:
            pages_str = "-"
            
        lines = []
        lines.append(f"**พบข้อผิดพลาดในหน้า:** {pages_str}\n")
        lines.append("**รายงานการตรวจทานเอกสาร:**\n")

        for idx, issue in enumerate(filtered):
            wrong = issue["wrong_word"]
            correct = issue["correct_word"]
            reason = issue["reason"]
            page_code = issue.get("page_code") or f"หน้า {issue.get('page_hint', 1)}"
            full_header = issue.get("full_header") or ""
            pos_str = f"{full_header}" if full_header else f"{page_code}"
            
            # ดึงประโยคบริบทและไฮไลต์คำผิดด้วยสีเหลืองสไตล์ Word (ไม่มีเครื่องหมายก้ามปู [[]])
            hl_tag = f'<mark style="background-color: #ffff00; color: #000000; padding: 1px 5px; border-radius: 2px; font-weight: bold;">{wrong}</mark>'
            snippet = (issue.get("snippet") or "").strip()
            if wrong and wrong in snippet:
                highlighted_snippet = snippet.replace(wrong, hl_tag, 1)
            elif snippet:
                highlighted_snippet = snippet
            else:
                highlighted_snippet = hl_tag
            
            lines.append(f"o  **หน้า / ตำแหน่ง:** {pos_str}  ")
            lines.append(f"o  **ข้อความในเอกสาร:** ...{highlighted_snippet}...  ")
            lines.append(f"o  **จุดที่ผิด:** ❌ 🔴 <span style=\"color: #b91c1c; font-weight: bold;\">{wrong}</span>  ")
            lines.append(f"o  **แก้ไขเป็น:** ✅ 🟢 <span style=\"color: #15803d; font-weight: bold;\">{correct}</span>  ")
            lines.append(f"o  **เหตุผล/คำแนะนำ:** {reason}\n")

        report_text = "\n".join(lines)

        with st.container(height=650):
            st.markdown(report_text, unsafe_allow_html=True)


# ============================================================
# Empty State (Drag and Drop Zone ขนาดใหญ่จัดการเรียบร้อยแล้ว)
# ============================================================
elif not uploaded_file:
    pass
