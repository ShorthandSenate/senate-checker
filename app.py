# ============================================================
# app.py - Streamlit Web Application ระดับผู้บริหาร
# ระบบตรวจทานและรับรองความถูกต้อง รายงานการประชุมวุฒิสภา
# (Senate Meeting Report Verification & Quality Assurance System)
# สำนักงานเลขาธิการวุฒิสภา (The Secretariat of the Senate)
# 100% Deterministic Rule & Database Engine — ข้อมูลปลอดภัย ไม่รั่วไหล
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
import json
import importlib
import pandas as pd
import streamlit as st

from checker_engine import run_full_check
import config
from config import RULES_CONFIG, MAX_FILE_SIZE_MB
import database_manager as dm

# ป้องกันปัญหา Streamlit Cloud Hot-reload ค้างโมดูลเวอร์ชันเก่าใน memory
if not hasattr(dm, "get_senate_personnel"):
    try:
        dm = importlib.reload(dm)
    except Exception:
        pass

try:
    import senate_names_checker
    reload_senate_names_checker = getattr(senate_names_checker, "reload_senate_names_checker", lambda: None)
except Exception:
    reload_senate_names_checker = lambda: None


def safe_get_senate_personnel(active_only: bool = False) -> list:
    """ดึงรายชื่อ สว. และผู้บริหารอย่างปลอดภัย 100% ป้องกัน AttributeError บน Cloud"""
    global dm
    if not hasattr(dm, "get_senate_personnel"):
        try:
            dm = importlib.reload(dm)
        except Exception:
            pass
    if hasattr(dm, "get_senate_personnel"):
        try:
            return dm.get_senate_personnel(active_only=active_only)
        except Exception:
            pass
    # Fallback อ่านไฟล์ JSON โดยตรงหากโมดูลใน memory ยังไม่อัปเดต
    jpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "senate_personnel.json")
    if os.path.exists(jpath):
        try:
            with open(jpath, "r", encoding="utf-8") as f:
                data = json.load(f)
            if active_only:
                return [p for p in data if p.get("is_active", 1) == 1]
            return data
        except Exception:
            pass
    return []


def safe_add_or_update_senate_person(**kwargs) -> bool:
    """บันทึกรายชื่อ สว. หรือผู้บริหารอย่างปลอดภัย"""
    global dm
    if hasattr(dm, "add_or_update_senate_person"):
        try:
            return dm.add_or_update_senate_person(**kwargs)
        except Exception:
            pass
    return False


@st.cache_resource
def get_soft_chime_b64() -> str:
    """สร้างไฟล์เสียง WAV สังเคราะห์ 2 โน้ต (D5 -> A5) นุ่มนวล ละมุน สบายหู สำหรับแจ้งเตือนเมื่อตรวจเสร็จ"""
    sample_rate = 44100
    duration = 1.1
    num_samples = int(sample_rate * duration)
    buf = io.BytesIO()

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
        if t >= 0:
            env1 = math.exp(-5.5 * t)
            val += 0.22 * math.sin(2.0 * math.pi * 587.33 * t) * env1
            val += 0.06 * math.sin(2.0 * math.pi * 1174.66 * t) * env1
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
    """สร้าง HTML ไฮไลต์เปรียบเทียบจุดต่างระดับพยางค์/อักขระ"""
    if not wrong or not correct:
        return html.escape(str(wrong)), html.escape(str(correct))
    w_clusters = get_thai_clusters(str(wrong))
    c_clusters = get_thai_clusters(str(correct))
    sm = difflib.SequenceMatcher(None, w_clusters, c_clusters)
    
    if sm.ratio() < 0.2:
        w_html = f'<span style="background:#fee2e2;color:#991b1b;padding:2px 6px;border-radius:4px;font-weight:700;">{html.escape(str(wrong))}</span>'
        c_html = f'<span style="background:#dcfce7;color:#166534;padding:2px 6px;border-radius:4px;font-weight:700;">{html.escape(str(correct))}</span>'
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
                w_out.append(f'<span style="background:#fee2e2;color:#991b1b;padding:2px 5px;border-radius:4px;font-weight:800;border-bottom:2px solid #ef4444;">{w_part}</span>')
            if c_part:
                c_out.append(f'<span style="background:#dcfce7;color:#166534;padding:2px 5px;border-radius:4px;font-weight:800;border-bottom:2px solid #22c55e;">{c_part}</span>')
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
    page_title="ระบบตรวจรายงานการประชุมวุฒิสภา | The Secretariat of the Senate",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={
        "About": "ระบบตรวจทานและรับรองความถูกต้อง รายงานการประชุมวุฒิสภา v3.5 Enterprise\nสำนักงานเลขาธิการวุฒิสภา (The Secretariat of the Senate)\n100% Deterministic Engine — ปลอดภัย ไม่รั่วไหล รวดเร็ว แม่นยำ",
    },
)

# ============================================================
# Custom Executive CSS
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400&display=swap');

html, body, [class*="css"], .stMarkdown, .stText, h1, h2, h3, h4, p, label, input, textarea {
    font-family: 'Sarabun', 'Leelawadee UI', 'Tahoma', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* คืนค่าฟอนต์ไอคอนให้ Streamlit ป้องกันไอคอนเพี้ยน */
[data-testid*="stIcon"],
[class*="material-symbols"],
[class*="material-icons"],
.material-symbols-rounded,
.material-icons,
button span,
[data-testid="stFileUploadDropzone"] span {
    font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
}

/* Background & Body */
.stApp {
    background-color: #f8fafc;
}

/* Executive Header */
.executive-header {
    background: linear-gradient(135deg, #0a192f 0%, #172a45 40%, #1e3a8a 100%);
    color: white;
    padding: 2.2rem 2.8rem;
    border-radius: 18px;
    margin-bottom: 1.8rem;
    box-shadow: 0 10px 30px rgba(10, 25, 47, 0.28);
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-bottom: 4px solid #f59e0b;
    position: relative;
    overflow: hidden;
}
.executive-header::after {
    content: "";
    position: absolute;
    top: -50%;
    right: -10%;
    width: 350px;
    height: 350px;
    background: radial-gradient(circle, rgba(245, 158, 11, 0.12) 0%, rgba(255, 255, 255, 0) 70%);
    border-radius: 50%;
    pointer-events: none;
}
.executive-org {
    font-size: 0.95rem;
    font-weight: 600;
    letter-spacing: 1px;
    color: #fbbf24;
    text-transform: uppercase;
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 0.4rem;
}
.executive-title {
    font-size: 2.1rem;
    font-weight: 800;
    margin: 0;
    line-height: 1.3;
    letter-spacing: -0.5px;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.25);
}
.executive-subtitle {
    font-size: 1.02rem;
    font-weight: 400;
    color: #cbd5e1;
    margin: 0.5rem 0 0 0;
    line-height: 1.5;
}
.badge-bar {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 1.2rem;
}
.badge-pill {
    font-size: 0.8rem;
    font-weight: 600;
    padding: 5px 14px;
    border-radius: 20px;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(255, 255, 255, 0.12);
    backdrop-filter: blur(8px);
    border: 1px solid rgba(255, 255, 255, 0.15);
    color: #f1f5f9;
}
.badge-pill-live {
    background: rgba(16, 185, 129, 0.2);
    border-color: rgba(16, 185, 129, 0.4);
    color: #6ee7b7;
}
.badge-pill-gold {
    background: rgba(245, 158, 11, 0.2);
    border-color: rgba(245, 158, 11, 0.4);
    color: #fde68a;
}

/* Executive Metric Cards */
.metric-container {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 14px;
    margin-bottom: 1.8rem;
}
.metric-box {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.2rem 1.1rem;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.04);
    position: relative;
    overflow: hidden;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    border-top: 4px solid #1e3a8a;
}
.metric-box:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.08);
}
.metric-box .metric-title {
    font-size: 0.82rem;
    font-weight: 600;
    color: #64748b;
    margin-bottom: 0.4rem;
    display: flex;
    align-items: center;
    gap: 6px;
}
.metric-box .metric-val {
    font-size: 2.1rem;
    font-weight: 800;
    line-height: 1;
    color: #0f172a;
}
.metric-box .metric-desc {
    font-size: 0.76rem;
    color: #94a3b8;
    margin-top: 0.4rem;
}

/* Feature Pillars Grid (Landing Page) */
.pillar-card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.3rem;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03);
    height: 100%;
    transition: all 0.2s ease;
}
.pillar-card:hover {
    border-color: #cbd5e1;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.07);
}
.pillar-icon {
    font-size: 1.8rem;
    margin-bottom: 0.6rem;
}
.pillar-title {
    font-size: 1.05rem;
    font-weight: 700;
    color: #0f172a;
    margin-bottom: 0.35rem;
}
.pillar-desc {
    font-size: 0.87rem;
    color: #64748b;
    line-height: 1.5;
}

/* Polished Dropzone */
[data-testid="stFileUploadDropzone"] {
    min-height: 220px !important;
    border: 2px dashed #3b82f6 !important;
    background: linear-gradient(180deg, #f8faff 0%, #f0f4ff 100%) !important;
    border-radius: 16px !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    align-items: center !important;
    padding: 2.2rem 1.5rem !important;
    margin-top: 0.5rem !important;
    margin-bottom: 1.5rem !important;
    cursor: pointer !important;
    transition: all 0.25s ease-in-out !important;
    box-shadow: inset 0 2px 8px rgba(59, 130, 246, 0.03) !important;
}
[data-testid="stFileUploadDropzone"]:hover {
    border-color: #1d4ed8 !important;
    background: #eef2ff !important;
    box-shadow: 0 8px 24px rgba(37, 99, 235, 0.12) !important;
}
[data-testid="stFileUploadDropzone"] button {
    margin-top: 10px !important;
    padding: 0.55rem 1.8rem !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}

/* Issue Card (Executive View) */
.executive-issue-card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
    border-left: 6px solid #3b82f6;
    transition: all 0.2s ease;
}
.executive-issue-card:hover {
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
}
.card-header-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px;
    padding-bottom: 0.8rem;
    margin-bottom: 1rem;
    border-bottom: 1px solid #f1f5f9;
}
.card-badges {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
}
.tag-badge {
    padding: 4px 10px;
    border-radius: 6px;
    font-size: 0.8rem;
    font-weight: 700;
    display: inline-flex;
    align-items: center;
    gap: 5px;
}
.tag-page {
    background: #e0e7ff;
    color: #3730a3;
    border: 1px solid #c7d2fe;
}
.tag-para {
    background: #f1f5f9;
    color: #475569;
    border: 1px solid #e2e8f0;
}
.diff-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 14px;
    margin: 0.9rem 0;
}
@media (max-width: 768px) {
    .diff-grid { grid-template-columns: 1fr; }
}
.diff-col {
    padding: 0.9rem 1.1rem;
    border-radius: 10px;
    font-size: 0.92rem;
}
.diff-wrong {
    background: #fff1f2;
    border: 1px solid #fecdd3;
    border-left: 4px solid #e11d48;
}
.diff-correct {
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    border-left: 4px solid #16a34a;
}
.diff-label {
    font-size: 0.74rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 0.35rem;
}
.diff-text {
    font-size: 1.08rem;
    font-weight: 700;
    word-break: break-word;
    white-space: pre-wrap;
}
.snippet-preview-box {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    margin-top: 0.6rem;
    font-size: 0.9rem;
    line-height: 1.8;
    color: #334155;
    word-break: break-word;
}
.reason-box {
    font-size: 0.88rem;
    color: #475569;
    margin-top: 0.7rem;
    padding-top: 0.7rem;
    border-top: 1px dashed #e2e8f0;
    display: flex;
    align-items: baseline;
    gap: 8px;
}

/* Hide footer */
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# Client-side Keepalive Heartbeat
# ป้องกัน Streamlit หลับเวลาเปิดแท็บค้างไว้
# ============================================================
st.markdown("""
<script>
(function() {
    if (window._stKeepAliveActive) return;
    window._stKeepAliveActive = true;
    var PING_INTERVAL_MS = 2.5 * 60 * 1000; // 2.5 นาที

    function keepAlive() {
        try {
            fetch(window.location.origin + '/_stcore/health', {
                method: 'GET',
                cache: 'no-store'
            }).catch(function() {});
        } catch(e) {}
    }

    setInterval(keepAlive, PING_INTERVAL_MS);
    document.addEventListener('visibilitychange', function() {
        if (!document.hidden) keepAlive();
    });
    window.addEventListener('focus', keepAlive);
})();
</script>
""", unsafe_allow_html=True)

# ============================================================
# Executive Header
# ============================================================
st.markdown("""
<div class="executive-header">
    <div class="executive-org">
        <span>🏛️</span>
        <span>สำนักงานเลขาธิการวุฒิสภา • The Secretariat of the Senate</span>
    </div>
    <h1 class="executive-title">ระบบตรวจสอบและรับรองความถูกต้อง รายงานการประชุมวุฒิสภา</h1>
    <p class="executive-subtitle">Senate Meeting Report Verification & Quality Assurance System • 100% Deterministic Engine</p>
    <div class="badge-bar">
        <span class="badge-pill badge-pill-live">
            <span style="font-size:0.6rem;">🟢</span> เซิร์ฟเวอร์พร้อมทำงาน 24/7 (Protected)
        </span>
        <span class="badge-pill badge-pill-gold">
            <span>🏛️</span> ทำเนียบ สว. & ผู้บริหาร ๒๑๕ ท่าน (กฎ ๒ เคาะ)
        </span>
        <span class="badge-pill">
            <span>📚</span> ฐานข้อมูลคำทับศัพท์ทางการ ๑,๕๖๑ คำ
        </span>
        <span class="badge-pill">
            <span>🔒</span> ความปลอดภัย 100% ข้อมูลไม่รั่วไหล (Zero Data Leak)
        </span>
        <span class="badge-pill">
            <span>⚡</span> v3.5 Enterprise Master (อัปเดต 6 ต.ค. 2569)
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# Sidebar: ทำเนียบ สว. และผู้บริหาร
# ============================================================
with st.sidebar:
    st.markdown("### 🏛️ ทำเนียบ สว. และผู้บริหาร")
    personnel_list = safe_get_senate_personnel(active_only=False)
    active_count = sum(1 for p in personnel_list if p.get("is_active", 1) == 1)
    senator_count = sum(1 for p in personnel_list if p.get("person_type") == "senator" and p.get("is_active", 1) == 1)
    exec_count = sum(1 for p in personnel_list if p.get("person_type") == "executive" and p.get("is_active", 1) == 1)

    st.success(f"👥 **กำลังปฏิบัติหน้าที่:** **{active_count} ท่าน**\n\n• สมาชิกวุฒิสภา: **{senator_count} ท่าน**\n• ผู้บริหารสำนักงานฯ: **{exec_count} ท่าน**")

    with st.expander("🔍 ค้นหารายชื่อในทำเนียบ", expanded=False):
        search_q = st.text_input("ค้นหาชื่อ/สกุล", key="search_person_q")
        matched = []
        for p in personnel_list:
            full = p.get("full_name_official", "")
            if not search_q or search_q.strip() in full:
                status_icon = "🟢" if p.get("is_active", 1) == 1 else "🔴 (พ้นตำแหน่ง)"
                matched.append(f"{status_icon} **{full}** — *{p.get('role', '')}*")
        if matched:
            st.markdown("\n\n".join(matched[:15]))
            if len(matched) > 15:
                st.caption(f"... และอีก {len(matched) - 15} ท่าน")
        else:
            st.caption("ไม่พบรายชื่อที่ค้นหา")

    with st.expander("➕ เพิ่ม / แก้ไขรายชื่อ (อัปเดตระบบ)", expanded=False):
        st.caption("ใช้เมื่อมีการเปลี่ยนคำนำหน้า ชื่อ-นามสกุล หรือมีผู้เข้ารับตำแหน่งใหม่/ลาออก")
        p_type = st.selectbox("ประเภทบุคลากร", ["สมาชิกวุฒิสภา", "ผู้บริหารสำนักงานฯ"], key="add_p_type")
        p_type_val = "senator" if p_type == "สมาชิกวุฒิสภา" else "executive"

        col_t1, col_t2 = st.columns([1, 2])
        with col_t1:
            p_title = st.text_input("คำนำหน้า/ยศ", placeholder="นาย / พลเอก", key="add_p_title")
        with col_t2:
            p_first = st.text_input("ชื่อตัว", placeholder="ชื่อ", key="add_p_first")

        p_last = st.text_input("นามสกุล", placeholder="นามสกุล", key="add_p_last")
        default_role = "สมาชิกวุฒิสภา" if p_type == "สมาชิกวุฒิสภา" else "รองเลขาธิการวุฒิสภา"
        p_role = st.text_input("ตำแหน่งทางการ", value=default_role, key="add_p_role")
        p_status = st.radio("สถานะการดำรงตำแหน่ง", ["ปฏิบัติหน้าที่ (Active)", "พ้นจากตำแหน่ง/ลาออก (Inactive)"], index=0, key="add_p_status")
        is_active_val = 1 if "Active" in p_status else 0

        if st.button("💾 บันทึกและอัปเดตระบบทันที", type="primary", use_container_width=True):
            if not p_first or not p_last:
                st.error("กรุณาระบุชื่อและนามสกุลให้ครบถ้วน")
            else:
                ok = safe_add_or_update_senate_person(
                    person_type=p_type_val,
                    title=p_title,
                    first_name=p_first,
                    last_name=p_last,
                    role=p_role,
                    is_active=is_active_val,
                )
                if ok:
                    reload_senate_names_checker()
                    st.success(f"✅ บันทึกข้อมูล '{p_title}{p_first}  {p_last}' เรียบร้อยแล้ว ระบบพร้อมใช้ตรวจทันที!")
                    st.rerun()
                else:
                    st.error("เกิดข้อผิดพลาดในการบันทึกข้อมูล")

    st.divider()
    st.markdown("### 💡 การใช้งาน & เทคนิคค้นหา")
    st.info(
        "**เทคนิคการทำงานกับไฟล์ Word:**\n\n"
        "1. ในผลการตรวจ ให้ดูข้อความใน **กล่องข้อความแวดล้อม**\n"
        "2. ลากแถบคลุมหรือก๊อปปี้คำนั้น\n"
        "3. สลับไปที่โปรแกรม Word แล้วกด **Ctrl + F**\n"
        "4. วางข้อความลงในช่องค้นหา จะกระโดดไปยังตำแหน่งที่ต้องแก้ไขทันที"
    )

# ============================================================
# Session State Initialization
# ============================================================
for key in ["check_results", "paragraphs", "last_filename"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ============================================================
# File Upload Area
# ============================================================
st.markdown("### 📂 อัปโหลดเอกสารรายงานการประชุมวุฒิสภา")
uploaded_file = st.file_uploader(
    "เลือกไฟล์ Word (.docx) หรือลากไฟล์มาวางในพื้นที่นี้ (รองรับ 200+ หน้า ขนาดสูงสุด 100 MB)",
    type=["docx"],
    help=f"ขนาดไฟล์สูงสุด {MAX_FILE_SIZE_MB} MB",
    label_visibility="collapsed",
)

# ============================================================
# Run Controls & Execution
# ============================================================
if uploaded_file:
    file_size_mb = uploaded_file.size / (1024 * 1024)
    if file_size_mb > MAX_FILE_SIZE_MB:
        st.error(f"❌ ไฟล์มีขนาดใหญ่เกินกำหนด ({file_size_mb:.1f} MB > {MAX_FILE_SIZE_MB} MB)")
        st.stop()

    c_btn1, c_btn2, c_sp = st.columns([2, 1, 3])
    with c_btn1:
        run_btn = st.button("🚀 เริ่มการตรวจสอบเอกสารทันที", type="primary", use_container_width=True)
    with c_btn2:
        if st.button("🗑️ ล้างผลการตรวจ", use_container_width=True):
            st.session_state.check_results = None
            st.session_state.paragraphs = None
            st.session_state.last_filename = None
            st.rerun()

    if run_btn:
        file_bytes = uploaded_file.read()
        st.session_state.last_filename = uploaded_file.name

        progress_bar = st.progress(0, text="กำลังเตรียมกระบวนการตรวจสอบ...")
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
            progress_bar.progress(1.0, text=f"✅ การตรวจสอบเสร็จสมบูรณ์ใน {elapsed:.2f} วินาที")
            status_text.success(
                f"✅ ตรวจสอบครบทุกย่อหน้า 100% ({len(paragraphs):,} ย่อหน้า) | "
                f"พบประเด็นข้อสังเกต {len(issues):,} จุด | "
                f"ใช้เวลาประมวลผล {elapsed:.2f} วินาที"
            )

            # เล่นเสียงสังเคราะห์แจ้งเตือนอย่างละมุน
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
            st.error(f"❌ เกิดข้อผิดพลาดระหว่างการตรวจสอบ: {err}")
            logging.exception("run_full_check error")

# ============================================================
# Empty State: Executive Presentation (เมื่อยังไม่มีไฟล์)
# ============================================================
elif not uploaded_file and st.session_state.check_results is None:
    st.markdown("""
    <div style="margin-top: 1rem; margin-bottom: 2rem;">
        <h4 style="color:#1e3a8a; font-weight:700; margin-bottom:1rem; display:flex; align-items:center; gap:8px;">
            <span>🛡️</span> เสาหลักการตรวจสอบความถูกต้อง ๗ ประการ (Quality Assurance Pillars)
        </h4>
        <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap:16px;">
            <div class="pillar-card">
                <div class="pillar-icon">🏛️</div>
                <div class="pillar-title">๑. ทำเนียบ สว. และผู้บริหาร (๒๑๕ ท่าน)</div>
                <div class="pillar-desc">ตรวจสอบคำนำหน้า ยศ การสะกดชื่อ-นามสกุล และบังคับใช้กฎ <b>เว้นวรรค ๒ เคาะ</b> ระหว่างชื่อตัวและนามสกุลตามระเบียบสารบรรณรัฐสภา 100%</div>
            </div>
            <div class="pillar-card">
                <div class="pillar-icon">📚</div>
                <div class="pillar-title">๒. คำทับศัพท์และศัพท์บัญญัติ (๑,๕๖๑ คำ)</div>
                <div class="pillar-desc">เทียบเคียงกับคลังคำศัพท์ทางการของวุฒิสภาและราชบัณฑิตยสภา ตรวจจับคำภาษาอังกฤษที่ควรใช้คำไทย และคำทับศัพท์ที่สะกดผิด</div>
            </div>
            <div class="pillar-card">
                <div class="pillar-icon">📄</div>
                <div class="pillar-title">๓. หัวแผ่นกระดาษและระบุหน้าแม่นยำ</div>
                <div class="pillar-desc">ตรวจจับรหัสหัวแผ่นกระดาษ (เช่น <i>๒๙/๑</i>) และเลขย่อหน้าอย่างละเอียด ช่วยให้ค้นหาจุดผิดในเอกสารต้นฉบับได้ใน ๒ วินาที</div>
            </div>
            <div class="pillar-card">
                <div class="pillar-icon">📏</div>
                <div class="pillar-title">๔. กฎวงเล็บภาษาอังกฤษซ้ำ</div>
                <div class="pillar-desc">อนุญาตให้ใส่วงเล็บภาษาอังกฤษขยายความได้เฉพาะครั้งแรกที่คำนั้นปรากฏในรายงานเท่านั้น หากพบในย่อหน้าถัดไปจะแจ้งเตือนให้ตัดออก</div>
            </div>
            <div class="pillar-card">
                <div class="pillar-icon">🔤</div>
                <div class="pillar-title">๕. คำสะกดผิดทางการและคำสลับพยัญชนะ</div>
                <div class="pillar-desc">ตรวจจับคำผิดยอดนิยมในรายงานการประชุม เช่น <i>สัมมนา, ผูกพัน, สังเกต, ลายเซ็น, ปาฐกถา</i> ตามพจนานุกรมฉบับราชบัณฑิตยสถาน</div>
            </div>
            <div class="pillar-card">
                <div class="pillar-icon">⚙️</div>
                <div class="pillar-title">๖. การใช้ไม้ยมก (ๆ) และเครื่องหมายวรรคตอน</div>
                <div class="pillar-desc">กวดขันการเว้นวรรคหน้าและหลังไม้ยมก และเครื่องหมายวรรคตอนตามหลักไวยากรณ์ทางการอย่างถูกต้องและเป็นระเบียบ</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# Results Display Dashboard
# ============================================================
if st.session_state.check_results is not None:
    issues     = st.session_state.check_results
    paragraphs = st.session_state.paragraphs
    filename   = st.session_state.last_filename or "รายงานการประชุม.docx"

    st.markdown("---")
    st.markdown(f"## 📊 แดชบอร์ดสรุปผลการตรวจสอบ — `{filename}`")

    # Group counts
    rule_counts = {}
    for iss in issues:
        lbl = iss["rule_label"]
        rule_counts[lbl] = rule_counts.get(lbl, 0) + 1

    color_map = {rc["label"]: rc["color"] for rc in RULES_CONFIG.values()}

    # KPI Metric Cards
    total_issues = len(issues)
    senate_issues = rule_counts.get("รายนาม สว./ผู้บริหาร (และกฎ ๒ เคาะ)", 0)
    vocab_issues = rule_counts.get("คำทับศัพท์", 0) + rule_counts.get("ศัพท์บัญญัติ", 0)
    spell_issues = rule_counts.get("คำผิดทั่วไป", 0) + rule_counts.get("การใช้ไม้ยมก (ๆ)", 0)
    paren_issues = rule_counts.get("วงเล็บซ้ำ", 0)

    st.markdown(f"""
    <div class="metric-container">
        <div class="metric-box" style="border-top-color:#1e3a8a;">
            <div class="metric-title">📋 ย่อหน้าที่ตรวจทานทั้งหมด</div>
            <div class="metric-val" style="color:#1e3a8a;">{len(paragraphs):,}</div>
            <div class="metric-desc">สแกนครบถ้วน 100%</div>
        </div>
        <div class="metric-box" style="border-top-color:#e11d48;">
            <div class="metric-title">⚠️ ข้อสังเกตที่พบทั้งหมด</div>
            <div class="metric-val" style="color:#e11d48;">{total_issues:,}</div>
            <div class="metric-desc">จุดที่ต้องแก้ไข/ปรับปรุง</div>
        </div>
        <div class="metric-box" style="border-top-color:#8b5cf6;">
            <div class="metric-title">🏛️ นาม สว. & ๒ เคาะ</div>
            <div class="metric-val" style="color:#8b5cf6;">{senate_issues:,}</div>
            <div class="metric-desc">รายนามและวรรคตอน ๒ เคาะ</div>
        </div>
        <div class="metric-box" style="border-top-color:#10b981;">
            <div class="metric-title">📚 คำทับศัพท์ & ศัพท์บัญญัติ</div>
            <div class="metric-val" style="color:#10b981;">{vocab_issues:,}</div>
            <div class="metric-desc">ตามมติและราชบัณฑิตฯ</div>
        </div>
        <div class="metric-box" style="border-top-color:#f59e0b;">
            <div class="metric-title">🔤 คำสะกดผิด & ไม้ยมก</div>
            <div class="metric-val" style="color:#f59e0b;">{spell_issues:,}</div>
            <div class="metric-desc">ไวยากรณ์และคำสลับพยัญชนะ</div>
        </div>
        <div class="metric-box" style="border-top-color:#64748b;">
            <div class="metric-title">📏 วงเล็บภาษาอังกฤษซ้ำ</div>
            <div class="metric-val" style="color:#64748b;">{paren_issues:,}</div>
            <div class="metric-desc">แนะนำให้ตัดวงเล็บออก</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not issues:
        st.success("🎉 **ยอดเยี่ยมมาก! การตรวจทานเอกสารเสร็จสมบูรณ์ 100% ไม่พบข้อผิดพลาดตามหลักเกณฑ์ที่กำหนดไว้เลยครับ**")
        filtered = []
    else:
        # Filters Bar
        st.markdown("### 🔎 กรองผลลัพธ์และค้นหาข้อมูล")
        fc1, fc2, fc3 = st.columns([2, 2, 2])
        with fc1:
            all_labels = sorted(set(i["rule_label"] for i in issues))
            filter_rule = st.multiselect("📌 ประเภทปัญหา", all_labels, default=all_labels)
        with fc2:
            page_codes_ordered = []
            for i in issues:
                pc = i.get("page_code") or f"~หน้า {i.get('page_hint', 1)}"
                if pc not in page_codes_ordered:
                    page_codes_ordered.append(pc)
            filter_page = st.selectbox("📄 กรองตามหัวแผ่นกระดาษ", ["ทุกหน้า (ทั้งหมด)"] + page_codes_ordered)
        with fc3:
            search_word = st.text_input("🔍 ค้นหาคำผิด / คำถูก / เนื้อหา", placeholder="พิมพ์คำที่ต้องการค้นหา...")

        fc4, fc5 = st.columns([3, 1])
        with fc4:
            max_para = max((i["para_index"] for i in issues), default=1)
            if max_para > 1:
                para_range = st.slider("ช่วงย่อหน้า", 1, max_para, (1, max_para))
            else:
                para_range = (1, 1)
        with fc5:
            sort_by = st.selectbox("เรียงตาม", ["ย่อหน้า", "หัวแผ่น / หน้าที่", "ประเภท"])

        # Apply Filters
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

        st.caption(f"แสดงผล **{len(filtered):,} รายการ** (จากข้อสังเกตทั้งหมด {len(issues):,} รายการ)")

        # Export & Actions
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

            col_exp1, col_exp2 = st.columns([2, 4])
            with col_exp1:
                st.download_button(
                    "⬇️ ดาวน์โหลดรายงาน Excel (.csv ภาษาไทยแท้)",
                    data=csv_bytes,
                    file_name=f"รายงานผลตรวจ_{filename.replace('.docx','')}.csv",
                    mime="text/csv",
                    type="primary",
                    use_container_width=True,
                )

        st.markdown("---")

        # View Mode Tabs: Executive Cards vs Official Text Report vs Data Table
        tab_cards, tab_text, tab_table = st.tabs([
            f"📑 การ์ดรายงานวิเคราะห์ ({len(filtered)})",
            "📋 รายงานข้อความทางการ (พร้อมคัดลอก)",
            "📊 ตารางข้อมูลสรุป (Data Table)",
        ])

        with tab_cards:
            if not filtered:
                st.info("ℹ️ ไม่พบประเด็นตามเงื่อนไขการกรองที่เลือก")
            else:
                for idx, issue in enumerate(filtered, 1):
                    wrong = issue["wrong_word"]
                    correct = issue["correct_word"]
                    reason = issue["reason"]
                    rule_lbl = issue["rule_label"]
                    badge_color = color_map.get(rule_lbl, "#3b82f6")
                    page_code = issue.get("page_code") or f"หน้า {issue.get('page_hint', 1)}"
                    full_header = issue.get("full_header") or ""
                    pos_str = f"{full_header}" if full_header else f"{page_code}"
                    snippet = (issue.get("snippet") or "").strip()

                    wrong_display = wrong.replace("  ", "&nbsp;&nbsp;")
                    correct_display = correct.replace("  ", "&nbsp;&nbsp;")
                    reason_display = reason.replace("  ", "&nbsp;&nbsp;")

                    hl_tag = f'<mark style="background-color: #fef08a; color: #854d0e; padding: 2px 6px; border-radius: 4px; font-weight: 800; border-bottom: 2px solid #ef4444; white-space: pre-wrap;">{wrong_display}</mark>'
                    if wrong and wrong in snippet:
                        highlighted_snippet = snippet.replace(wrong, hl_tag, 1)
                    elif snippet:
                        highlighted_snippet = snippet
                    else:
                        highlighted_snippet = hl_tag
                    highlighted_snippet = highlighted_snippet.replace("  ", "&nbsp;&nbsp;")

                    st.markdown(f"""
                    <div class="executive-issue-card" style="border-left-color: {badge_color};">
                        <div class="card-header-bar">
                            <div class="card-badges">
                                <span class="tag-badge" style="background:{badge_color}18; color:{badge_color}; border:1px solid {badge_color}35;">
                                    📌 {rule_lbl}
                                </span>
                                <span class="tag-badge tag-page">
                                    📄 {pos_str}
                                </span>
                                <span class="tag-badge tag-para">
                                    ย่อหน้าที่ {issue['para_index']}
                                </span>
                            </div>
                            <span style="font-size:0.82rem; font-weight:700; color:#94a3b8;">
                                ลำดับที่ #{idx}
                            </span>
                        </div>
                        <div class="diff-grid">
                            <div class="diff-col diff-wrong">
                                <div class="diff-label" style="color:#e11d48;">❌ คำที่ปรากฏในเอกสาร (คำเดิม)</div>
                                <div class="diff-text" style="color:#be123c;">{wrong_display}</div>
                            </div>
                            <div class="diff-col diff-correct">
                                <div class="diff-label" style="color:#16a34a;">✅ แนะนำให้แก้ไขเป็น (คำที่ถูกต้อง)</div>
                                <div class="diff-text" style="color:#15803d;">{correct_display}</div>
                            </div>
                        </div>
                        <div class="snippet-preview-box">
                            <div style="font-size:0.75rem; font-weight:700; color:#64748b; margin-bottom:4px;">บริบทแวดล้อมในเอกสาร:</div>
                            <span style="white-space: pre-wrap;">...{highlighted_snippet}...</span>
                        </div>
                        <div class="reason-box">
                            <span style="font-weight:700; color:#1e3a8a;">💡 ระเบียบ/เหตุผล:</span>
                            <span>{reason_display}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

        with tab_text:
            if not filtered:
                st.info("ℹ️ ไม่พบประเด็นตามเงื่อนไขการกรอง")
            else:
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
                lines.append(f"**รายงานการตรวจทานเอกสาร: {filename}**\n")
                lines.append(f"**พบข้อสังเกตในหน้า:** {pages_str}\n")
                lines.append(f"**จำนวนประเด็นที่พบ:** {len(filtered):,} รายการ\n")
                lines.append("──────────────────────────────────────────\n")

                for idx, issue in enumerate(filtered, 1):
                    wrong = issue["wrong_word"]
                    correct = issue["correct_word"]
                    reason = issue["reason"]
                    page_code = issue.get("page_code") or f"หน้า {issue.get('page_hint', 1)}"
                    full_header = issue.get("full_header") or ""
                    pos_str = f"{full_header}" if full_header else f"{page_code}"

                    wrong_display = wrong.replace("  ", "&nbsp;&nbsp;")
                    correct_display = correct.replace("  ", "&nbsp;&nbsp;")
                    reason_display = reason.replace("  ", "&nbsp;&nbsp;")

                    hl_tag = f'<mark style="background-color: #fef08a; color: #854d0e; padding: 1px 4px; font-weight: bold; text-decoration: underline wavy red; white-space: pre-wrap;">{wrong_display}</mark>'
                    snippet = (issue.get("snippet") or "").strip()
                    if wrong and wrong in snippet:
                        highlighted_snippet = snippet.replace(wrong, hl_tag, 1)
                    elif snippet:
                        highlighted_snippet = snippet
                    else:
                        highlighted_snippet = hl_tag
                    highlighted_snippet = highlighted_snippet.replace("  ", "&nbsp;&nbsp;")

                    lines.append(f"• **ลำดับที่ {idx}** | **หน้า / ตำแหน่ง:** {pos_str} (ย่อหน้าที่ {issue['para_index']})")
                    lines.append(f'• **ข้อความในเอกสาร:** <span style="white-space: pre-wrap;">...{highlighted_snippet}...</span>')
                    lines.append(f'• **จุดที่ผิด:** ❌ 🔴 <span style="color: #b91c1c; font-weight: bold; text-decoration: underline wavy red; white-space: pre-wrap;">{wrong_display}</span>')
                    lines.append(f'• **แก้ไขเป็น:** ✅ 🟢 <span style="color: #15803d; font-weight: bold; white-space: pre-wrap;">{correct_display}</span>')
                    lines.append(f"• **เหตุผล/คำแนะนำ:** {reason_display}\n")

                report_text = "\n".join(lines)
                with st.container(height=650):
                    st.markdown(report_text, unsafe_allow_html=True)

        with tab_table:
            if not filtered:
                st.info("ℹ️ ไม่มีข้อมูล")
            else:
                table_data = [{
                    "ลำดับ": idx + 1,
                    "หน้า/ตำแหน่ง": i.get("page_code", f"~หน้า {i.get('page_hint', 1)}"),
                    "ย่อหน้า": i["para_index"],
                    "ประเภทกฎ": i["rule_label"],
                    "คำผิด": i["wrong_word"],
                    "คำที่ถูกต้อง": i["correct_word"],
                    "เหตุผล/คำแนะนำ": i["reason"],
                } for idx, i in enumerate(filtered)]
                st.dataframe(pd.DataFrame(table_data), use_container_width=True, height=500)
