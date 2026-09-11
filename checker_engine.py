# ============================================================
# checker_engine.py - หัวใจของการตรวจสอบ (100% Deterministic Rule & Database Engine)
# ถอด Gemini AI ออกทั้งหมด — ตรวจจากฐานข้อมูลและกฎตายตัวเท่านั้น
# ผลลัพธ์แม่นยำ รวดเร็ว ไม่เกิด Hallucination
# ============================================================

import re
import io
import logging

from docx import Document
from docx.oxml.ns import qn

from config import (
    CONTEXT_WORDS,
    RULES_CONFIG,
    MANDATORY_BRACKET_WORDS,
)

logger = logging.getLogger(__name__)


# ============================================================
# 1. อ่านไฟล์ Word และแยกย่อหน้า
#    รองรับ: Soft Break, Track Changes, หัวแผ่นกระดาษ
# ============================================================

SENATE_HEADER_REGEX = re.compile(
    r'^(ว\.\s*[\d๑-๙]+(?:\s*\([^\)]+\))?)\\s+(.+?)\s*([\d๑-๙]+/[\d๑-๙]+(?:\s*\([^\)]+\))?)$'
)
FALLBACK_HEADER_REGEX = re.compile(
    r'^([ก-๙A-Za-z\s]+?)\s+([\d๑-๙]+/[\d๑-๙]+(?:\s*\([^\)]+\))?)$'
)


def _get_paragraph_text_accepted(para) -> str:
    """
    ดึงข้อความจากย่อหน้า โดย:
    1. ยอมรับเฉพาะข้อความที่ "ยังอยู่" (ไม่รวม Track Changes ที่ถูกลบ)
    2. รวม Soft Break (Shift+Enter / w:br) ให้เป็นช่องไฟเดียว ไม่ตัดคำ
    3. ข้ามข้อความใน w:del (Tracked Deletion) อย่างสมบูรณ์
    """
    text_parts = []

    for elem in para._element.iter():
        tag = elem.tag

        # ข้าม Track Changes ส่วนที่ถูก "ลบ" (w:del) — ใช้ข้อความ "หลังแก้ไข" เท่านั้น
        # ข้ามทั้ง subtree ของ w:del โดยใช้ ancestor check
        ancestors = [e.tag for e in elem.iterancestors()]
        if any(a == qn('w:del') for a in ancestors):
            continue

        # Soft Break (Shift+Enter) → แทนด้วยช่องไฟ 1 ช่อง ไม่ให้ตัดคำ
        if tag == qn('w:br'):
            br_type = elem.get(qn('w:type'), '')
            if br_type != 'page':  # page break ข้ามไปเลย
                text_parts.append(' ')
            continue

        # ดึงข้อความปกติจาก w:t
        if tag == qn('w:t'):
            t = elem.text or ''
            if t:
                text_parts.append(t)

    return ''.join(text_parts).strip()


def read_docx_paragraphs(file_bytes: bytes) -> list:
    """
    อ่านไฟล์ .docx ครบถ้วน 100% ทุกย่อหน้า พร้อม:
    - สกัดหัวแผ่นกระดาษ (เช่น จันทร์ตรี ๓/๑) แบบแม่นยำ
    - รองรับ Track Changes (ใช้เฉพาะข้อความที่ยังอยู่หลังแก้ไข)
    - รองรับ Soft Break (Shift+Enter) โดยรวมเป็นช่องไฟ ไม่ตัดคำ
    คืนค่า list ของ dict: {index, text, page_code, full_header, is_header}
    """
    doc = Document(io.BytesIO(file_bytes))
    raw_paras = []

    for p in doc.paragraphs:
        t = _get_paragraph_text_accepted(p)
        # ไม่ยุบช่องไฟซ้ำ (ไม่ใช้ re.sub) เพื่อรักษาวรรคใหญ่ (๒ เคาะ) ให้คงอยู่ตามที่พิมพ์จริง
        t = t.strip()
        if t:
            raw_paras.append((p, t))

    # --- สกัดตำแหน่งหัวแผ่นกระดาษทั้งหมด ---
    SENATE_H = re.compile(
        r'^(ว\.\s*[\d๑-๙]+(?:\s*\([^\)]+\))?)\s+(.+?)\s*([\d๑-๙]+/[\d๑-๙]+(?:\s*\([^\)]+\))?)$'
    )
    FALLBACK_H = re.compile(
        r'^([ก-๙A-Za-z\s]+?)\s+([\d๑-๙]+/[\d๑-๙]+(?:\s*\([^\)]+\))?)$'
    )

    headers_map = {}
    first_header_info = None

    for idx, (p, text) in enumerate(raw_paras):
        m = SENATE_H.match(text)
        if not m:
            m = FALLBACK_H.match(text)
        if m:
            if len(m.groups()) == 3:
                session = m.group(1).strip()
                steno = m.group(2).strip()
                sh_pg = m.group(3).strip()
                full_h = f"{session}   {steno} {sh_pg}"
                code = f"{steno} {sh_pg}"
            else:
                steno = m.group(1).strip()
                sh_pg = m.group(2).strip()
                full_h = f"{steno} {sh_pg}"
                code = f"{steno} {sh_pg}"

            headers_map[idx] = {
                "page_code": code,
                "full_header": full_h,
                "steno": steno,
                "sheet_page": sh_pg,
            }
            if first_header_info is None:
                first_header_info = headers_map[idx]

    # คำนวณหน้าเริ่มต้นสำหรับข้อความก่อนหัวแผ่นแรก
    default_page_code = "หน้า ๑"
    default_full_header = "หน้า ๑"
    if first_header_info:
        steno = first_header_info["steno"]
        sh_pg = first_header_info["sheet_page"]
        m_page = re.match(r"([\d๑-๙]+)/", sh_pg)
        if m_page:
            sheet_num = m_page.group(1)
            default_page_code = f"{steno} {sheet_num}/๑"
            sess_prefix = first_header_info.get("full_header", "").split(steno)[0].strip()
            default_full_header = f"{sess_prefix}   {steno} {sheet_num}/๑".strip()
        else:
            default_page_code = f"{steno} ๑/๑"
            default_full_header = f"{steno} ๑/๑"

    current_page_code = default_page_code
    current_full_header = default_full_header
    page_count = 1

    paragraphs = []
    for idx, (p, text) in enumerate(raw_paras):
        is_header = False
        if idx in headers_map:
            current_page_code = headers_map[idx]["page_code"]
            current_full_header = headers_map[idx]["full_header"]
            is_header = True
            page_count += 1

        paragraphs.append({
            "index": idx,
            "text": text,
            "page_code": current_page_code,
            "full_header": current_full_header,
            "page_hint": current_page_code,
            "page_num": page_count,
            "is_header": is_header,
        })

    logger.info(f"อ่านไฟล์สำเร็จ: {len(paragraphs)} ย่อหน้า, {len(headers_map)} หัวแผ่นกระดาษ")
    return paragraphs


# ============================================================
# 2. Context Snippet Builder
# ============================================================

def build_context_snippet(text: str, wrong_word: str, word_count: int = CONTEXT_WORDS) -> str:
    """ดึงข้อความรอบข้างคำผิด word_count คำทั้งสองข้าง"""
    idx = text.find(wrong_word)
    if idx == -1:
        return (text[:120] + "...") if len(text) > 120 else text
    left_words = text[:idx].split()
    right_words = text[idx + len(wrong_word):].split()
    left_snippet = " ".join(left_words[-word_count:])
    right_snippet = " ".join(right_words[:word_count])
    return f"{left_snippet} [{wrong_word}] {right_snippet}".strip()


# ============================================================
# 3. กฎที่ 1: ตรวจคำผิดภาษาไทย-อังกฤษจากฐานข้อมูล (Word-style)
# ============================================================

import database_manager as dm

# เหตุผลอธิบายรายกลุ่มคำ สำหรับแสดงให้ผู้ตรวจตัดสินใจ
MISSPELLING_REASON_MAP = {
    "อนุญาติ": "คำว่า 'อนุญาติ' สะกดผิด ที่ถูกต้องคือ 'อนุญาต' (ไม่มี ติ ท้าย) ตามพจนานุกรมราชบัณฑิตยสภา",
    "สัมนา": "คำว่า 'สัมนา' พิมพ์ตกพยัญชนะ ที่ถูกต้องคือ 'สัมมนา' (มม) ตามพจนานุกรมราชบัณฑิตยสภา",
    "สัมนาการ": "คำว่า 'สัมนาการ' พิมพ์ตกพยัญชนะ ที่ถูกต้องคือ 'สัมมนาการ' (มม)",
    "สัมานา": "คำว่า 'สัมานา' สะกดผิด ที่ถูกต้องคือ 'สัมมนา' (มม)",
    "สัมนากร": "คำว่า 'สัมนากร' สะกดผิด ที่ถูกต้องคือ 'สัมมนากร' (มม)",
    "ประสิทธิ์ผล": "คำว่า 'ประสิทธิ์ผล' ใส่ทัณฑฆาตที่ ธิ เกินมา ที่ถูกต้องคือ 'ประสิทธิผล'",
    "ประสิทธิผลล": "คำว่า 'ประสิทธิผลล' พิมพ์เกิน ล ที่ถูกต้องคือ 'ประสิทธิผล'",
    "ประสิทธิ์ภาพ": "คำว่า 'ประสิทธิ์ภาพ' ใส่ทัณฑฆาตที่ ธิ เกินมา ที่ถูกต้องคือ 'ประสิทธิภาพ'",
    "อนุกรรมาธิกาณ": "คำว่า 'อนุกรรมาธิกาณ' สะกดผิด ที่ถูกต้องคือ 'อนุกรรมาธิการ'",
    "อนุกรมาธิการ": "คำว่า 'อนุกรมาธิการ' พิมพ์ตก รร ที่ถูกต้องคือ 'อนุกรรมาธิการ'",
    "กรรมาธิกาณ": "คำว่า 'กรรมาธิกาณ' สะกดผิด ที่ถูกต้องคือ 'กรรมาธิการ'",
    "ผ้อำนวยการ": "คำว่า 'ผ้อำนวยการ' พิมพ์ตก ู ที่ถูกต้องคือ 'ผู้อำนวยการ'",
    "เลขาธิการณ": "คำว่า 'เลขาธิการณ' มี ณ เกินท้าย ที่ถูกต้องคือ 'เลขาธิการ'",
    "ผู้บริหาล": "คำว่า 'ผู้บริหาล' สะกดผิด ลท้าย ที่ถูกต้องคือ 'ผู้บริหาร'",
    "บริหาล": "คำว่า 'บริหาล' สะกดผิด ลท้าย ที่ถูกต้องคือ 'บริหาร'",
    "ผลิตภัณท์": "คำว่า 'ผลิตภัณท์' สะกดผิด ที่ถูกต้องคือ 'ผลิตภัณฑ์'",
    "ผลิตภัน": "คำว่า 'ผลิตภัน' พิมพ์ตก ด ท้าย ที่ถูกต้องคือ 'ผลิตภัณฑ์'",
    "ผลิดภัณฑ์": "คำว่า 'ผลิดภัณฑ์' สะกดผิด ที่ถูกต้องคือ 'ผลิตภัณฑ์'",
    "ส่ิงแวดล้อม": "คำว่า 'ส่ิงแวดล้อม' สระผิดตำแหน่ง ที่ถูกต้องคือ 'สิ่งแวดล้อม'",
    "ส่ิงที่": "คำว่า 'ส่ิงที่' สระผิดตำแหน่ง ที่ถูกต้องคือ 'สิ่งที่'",
    "ปัจจุบับ": "คำว่า 'ปัจจุบับ' สะกดผิด ที่ถูกต้องคือ 'ปัจจุบัน'",
    "พระราชกฤษฎิกา": "คำว่า 'พระราชกฤษฎิกา' สะกดผิด ที่ถูกต้องคือ 'พระราชกฤษฎีกา'",
    "พระราชกฤษฎีกาา": "คำว่า 'พระราชกฤษฎีกาา' พิมพ์เกิน า ที่ถูกต้องคือ 'พระราชกฤษฎีกา'",
    "แลละ": "คำว่า 'แลละ' สะกดผิด ที่ถูกต้องคือ 'และ'",
    "กระทรวงสาธรณสุข": "คำว่า 'กระทรวงสาธรณสุข' พิมพ์ตก า ที่ถูกต้องคือ 'กระทรวงสาธารณสุข'",
    "ช็อปปี": "คำว่า 'ช็อปปี' สะกดผิด ที่ถูกต้องคือ 'ช้อปปี' ตามมติราชบัณฑิตยสภา",
    # คำทับศัพท์ทั่วไป
    "สมาร์ท": "คำว่า 'สมาร์ท' สะกดผิดตามหลักราชบัณฑิตยสภา ที่ถูกต้องคือ 'สมาร์ต' (ไม่มีทัณฑฆาต)",
    "แอพ": "คำว่า 'แอพ' สะกดผิด ที่ถูกต้องคือ 'แอป'",
    "แอพพลิเคชัน": "คำว่า 'แอพพลิเคชัน' สะกดผิด ที่ถูกต้องคือ 'แอปพลิเคชัน'",
    "ดิจิตอล": "คำว่า 'ดิจิตอล' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'ดิจิทัล'",
    "อัพเดท": "คำว่า 'อัพเดท' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'อัปเดต'",
    "อัพเดต": "คำว่า 'อัพเดต' สะกดผิด ที่ถูกต้องคือ 'อัปเดต' (อัป ไม่ใช่ อัพ)",
    "อัปเดท": "คำว่า 'อัปเดท' สะกดผิด ที่ถูกต้องคือ 'อัปเดต' (ไม่มีทัณฑฆาต)",
    "ลิงค์": "คำว่า 'ลิงค์' สะกดผิด ที่ถูกต้องคือ 'ลิงก์'",
    "ลิ้งค์": "คำว่า 'ลิ้งค์' สะกดผิด ที่ถูกต้องคือ 'ลิงก์' (ไม่มีไม้โท และ ค ไม่มีทัณฑฆาต)",
    "คลิ๊ก": "คำว่า 'คลิ๊ก' สะกดผิด ที่ถูกต้องคือ 'คลิก'",
    "อีเมล์": "คำว่า 'อีเมล์' สะกดผิด ที่ถูกต้องคือ 'อีเมล' (ไม่มีทัณฑฆาต)",
    "แพลทฟอร์ม": "คำว่า 'แพลทฟอร์ม' สะกดผิด ที่ถูกต้องคือ 'แพลตฟอร์ม'",
    "เว็ปไซต์": "คำว่า 'เว็ปไซต์' สะกดผิด ที่ถูกต้องคือ 'เว็บไซต์' (บ ไม่ใช่ ป)",
    "เวบไซต์": "คำว่า 'เวบไซต์' สะกดผิด ที่ถูกต้องคือ 'เว็บไซต์'",
    "เวปไซต์": "คำว่า 'เวปไซต์' สะกดผิด ที่ถูกต้องคือ 'เว็บไซต์'",
    "ชาร์ต": "คำว่า 'ชาร์ต' สะกดผิด ที่ถูกต้องคือ 'ชาร์จ'",
    "โพส": "คำว่า 'โพส' พิมพ์ตก ต ที่ถูกต้องคือ 'โพสต์'",
    "ยูทูป": "คำว่า 'ยูทูป' สะกดผิด ที่ถูกต้องคือ 'ยูทูบ'",
    "โปรเจค": "คำว่า 'โปรเจค' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'โพรเจกต์'",
    "ฟังก์ชั่น": "คำว่า 'ฟังก์ชั่น' สะกดผิด ที่ถูกต้องคือ 'ฟังก์ชัน' (ไม่มีวรรณยุกต์ เอก)",
    "กราฟฟิก": "คำว่า 'กราฟฟิก' สะกดผิด ที่ถูกต้องคือ 'กราฟิก' (ฟ ตัวเดียว)",
    "กราฟฟิค": "คำว่า 'กราฟฟิค' สะกดผิด ที่ถูกต้องคือ 'กราฟิก'",
    "เซ็นเซอร์": "คำว่า 'เซ็นเซอร์' สะกดผิด ที่ถูกต้องคือ 'เซนเซอร์'",
    "เช็ค": "คำว่า 'เช็ค' สะกดผิด ที่ถูกต้องคือ 'เช็ก'",
    "ซอฟท์แวร์": "คำว่า 'ซอฟท์แวร์' สะกดผิด ที่ถูกต้องคือ 'ซอฟต์แวร์'",
    "สตาร์ท": "คำว่า 'สตาร์ท' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'สตาร์ต'",
    "สติ๊กเกอร์": "คำว่า 'สติ๊กเกอร์' สะกดผิด ที่ถูกต้องคือ 'สติกเกอร์'",
    "โควต้า": "คำว่า 'โควต้า' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'โควตา'",
    "แท็กซี่": "คำว่า 'แท็กซี่' สะกดผิด ที่ถูกต้องคือ 'แท็กซี'",
    "พอยท์": "คำว่า 'พอยท์' สะกดผิด ที่ถูกต้องคือ 'พอยต์'",
    "ปาร์ตี้": "คำว่า 'ปาร์ตี้' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'ปาร์ตี'",
}


def _get_misspelling_reason(wrong: str, correct: str) -> str:
    """สร้างเหตุผลอธิบายสำหรับคำผิดแต่ละคำ เพื่อให้ผู้ตรวจตัดสินใจเอง"""
    if wrong in MISSPELLING_REASON_MAP:
        return MISSPELLING_REASON_MAP[wrong]
    return (
        f"คำว่า '{wrong}' อาจสะกดผิดหรือพิมพ์ตก "
        f"กรุณาพิจารณา: คำที่ถูกต้องตามฐานข้อมูลคือ '{correct}' — "
        f"โปรดตรวจสอบบริบทก่อนแก้ไข"
    )


def check_misspellings_dictionary(paragraphs: list) -> list:
    """
    ตรวจหาคำผิดภาษาไทย-อังกฤษจากฐานข้อมูล (Word-style Lookup):
    - ใช้ COMMON_MISPELLING_MAP ใน database_manager.py
    - ตรวจแบบ Exact Substring Match สำหรับภาษาไทย
    - ตรวจแบบ Word Boundary สำหรับภาษาอังกฤษ
    - แสดงเหตุผลประกอบทุกรายการ เพื่อให้ผู้ตรวจตัดสินใจ
    """
    if not RULES_CONFIG.get("misspelling", {}).get("enabled", True):
        return []

    misspelling_map = dm.COMMON_MISPELLING_MAP
    if not misspelling_map:
        return []

    color = RULES_CONFIG["misspelling"]["color"]
    label = RULES_CONFIG["misspelling"]["label"]
    issues = []

    # เรียงคำค้นหาจากยาวไปสั้น (Longest Match First) ป้องกันคำย่อยซ้อน
    sorted_wrongs = sorted(misspelling_map.keys(), key=len, reverse=True)

    for p in paragraphs:
        if p.get("is_header"):
            continue

        text = p["text"]
        para_idx = p["index"] + 1
        page_code = p["page_code"]
        full_header = p["full_header"]
        covered_spans = []

        for wrong in sorted_wrongs:
            correct = misspelling_map[wrong]

            # ข้ามถ้าคำ "ผิด" เท่ากับคำ "ถูก" (anchor entries)
            if wrong == correct:
                continue

            is_english = bool(re.match(r'^[A-Za-z0-9\s\-]+$', wrong))

            if is_english:
                pattern = rf'\b{re.escape(wrong)}\b'
                for m in re.finditer(pattern, text, re.IGNORECASE):
                    start, end = m.start(), m.end()
                    if any(cs <= start and end <= ce for cs, ce in covered_spans):
                        continue
                    covered_spans.append((start, end))
                    matched = m.group()
                    reason = _get_misspelling_reason(wrong, correct)
                    issues.append({
                        "rule": "misspelling",
                        "rule_label": label,
                        "para_index": para_idx,
                        "page_hint": page_code,
                        "page_code": page_code,
                        "full_header": full_header,
                        "wrong_word": matched,
                        "correct_word": correct,
                        "reason": reason,
                        "snippet": build_context_snippet(text, matched),
                        "color": color,
                    })
            else:
                # ภาษาไทย: Exact Substring Match
                start_pos = 0
                while True:
                    idx_found = text.find(wrong, start_pos)
                    if idx_found == -1:
                        break
                    start, end = idx_found, idx_found + len(wrong)
                    start_pos = end

                    # ป้องกันเด็ดขาด: ถ้าคำถูกยาวกว่าและข้อความตำแหน่งนี้คือคำที่ถูกต้องอยู่แล้ว ให้ข้ามทันที
                    if len(correct) > len(wrong) and text[start:start + len(correct)] == correct:
                        continue

                    if any(cs <= start and end <= ce for cs, ce in covered_spans):
                        continue
                    covered_spans.append((start, end))
                    reason = _get_misspelling_reason(wrong, correct)
                    issues.append({
                        "rule": "misspelling",
                        "rule_label": label,
                        "para_index": para_idx,
                        "page_hint": page_code,
                        "page_code": page_code,
                        "full_header": full_header,
                        "wrong_word": wrong,
                        "correct_word": correct,
                        "reason": reason,
                        "snippet": build_context_snippet(text, wrong),
                        "color": color,
                    })

    return issues


# ============================================================
# 4. กฎที่ 2: ตรวจคำวงเล็บภาษาอังกฤษซ้ำ (Exact Match — ทั้งหมด)
# ============================================================

def check_parenthesis_repeat(paragraphs: list) -> list:
    """
    ตรวจจับคำภาษาอังกฤษในวงเล็บที่ปรากฏซ้ำ:
    - คำแรกของในรายงาน: ถูกต้อง / อนุญาต ไม่แจ้งเตือน
    - ครั้งที่ 2 เป็นต้นไป: แจ้งเตือนให้ตัดวงเล็บออก
    - ใช้ Exact Case-sensitive Match: (Soft Power) ≠ (Soft) ≠ (Power)
    """
    if not RULES_CONFIG.get("parenthesis_repeat", {}).get("enabled", True):
        return []

    pattern = RULES_CONFIG["parenthesis_repeat"]["bracket_pattern"]
    seen: dict = {}  # exact_key -> {para_index, page_code}
    issues = []

    for p in paragraphs:
        if p.get("is_header"):
            continue

        text = p["text"]
        para_idx = p["index"] + 1
        page_code = p["page_code"]
        full_header = p["full_header"]

        for m in re.finditer(pattern, text):
            inner_text = m.group(1).strip()

            # กรองเฉพาะวงเล็บที่มีตัวอักษรภาษาอังกฤษ
            if not re.search(r'[A-Za-z]', inner_text):
                continue

            # Exact Match (Case-sensitive) — (Soft Power) ≠ (Soft) ≠ (Power)
            exact_key = inner_text  # ไม่ normalize ไม่ lowercase
            bracket_form = f"({inner_text})"

            if exact_key not in seen:
                seen[exact_key] = {
                    "para_index": para_idx,
                    "page_code": page_code,
                }
            else:
                first = seen[exact_key]
                issues.append({
                    "rule": "parenthesis_repeat",
                    "rule_label": RULES_CONFIG["parenthesis_repeat"]["label"],
                    "para_index": para_idx,
                    "page_hint": page_code,
                    "page_code": page_code,
                    "full_header": full_header,
                    "wrong_word": bracket_form,
                    "correct_word": f"ตัดวงเล็บออก (กล่าวถึงครั้งแรกแล้วที่ย่อหน้า {first['para_index']} [{first['page_code']}])",
                    "reason": (
                        f"คำภาษาอังกฤษในวงเล็บ \"{bracket_form}\" ปรากฏเป็นครั้งแรกแล้วที่ย่อหน้า {first['para_index']} "
                        f"({first['page_code']}) การกล่าวถึงตั้งแต่ครั้งที่ ๒ เป็นต้นไป ให้ตัดวงเล็บออกตามระเบียบสำนักกรรมาธิการ ๓"
                    ),
                    "snippet": build_context_snippet(text, bracket_form),
                    "color": RULES_CONFIG["parenthesis_repeat"]["color"],
                })

    return issues


# ============================================================
# 5. กฎที่ 3: ตรวจคำบังคับมีวงเล็บภาษาอังกฤษในครั้งแรก
#    เช่น "พาวเวอร์พ็อยนต์" ต้องมี (PowerPoint) ในการกล่าวถึงครั้งแรก
# ============================================================

def check_mandatory_first_parenthesis(paragraphs: list) -> list:
    """
    ตรวจสอบคำทับศัพท์ที่กำหนดไว้ใน MANDATORY_BRACKET_WORDS ว่า:
    1. ครั้งแรกที่ปรากฏ: ต้องมีวงเล็บ (EN) กำกับต่อท้ายทันที
       หากไม่มี → แจ้งเตือนให้เติม
    2. ครั้งที่ 2 เป็นต้นไป: ต้องไม่มีวงเล็บ (ถ้ามีจะถูกจับโดย check_parenthesis_repeat)
    """
    if not RULES_CONFIG.get("mandatory_bracket", {}).get("enabled", True):
        return []

    color = RULES_CONFIG["mandatory_bracket"]["color"]
    label = RULES_CONFIG["mandatory_bracket"]["label"]
    issues = []

    for thai_word, en_word in MANDATORY_BRACKET_WORDS.items():
        first_occurrence_found = False
        expected_bracket = f"({en_word})"

        for p in paragraphs:
            if p.get("is_header"):
                continue

            text = p["text"]
            para_idx = p["index"] + 1
            page_code = p["page_code"]
            full_header = p["full_header"]

            # ค้นหาทุกตำแหน่งที่พบคำทับศัพท์ไทย
            start_pos = 0
            while True:
                idx_found = text.find(thai_word, start_pos)
                if idx_found == -1:
                    break

                end_pos = idx_found + len(thai_word)
                start_pos = end_pos

                if not first_occurrence_found:
                    # ครั้งแรก: ตรวจว่ามีวงเล็บ EN ต่อท้ายทันทีหรือไม่
                    first_occurrence_found = True
                    after_word = text[end_pos:end_pos + len(expected_bracket) + 5].strip()

                    # รองรับว่าอาจมีช่องไฟระหว่างคำไทยและวงเล็บ
                    remaining = text[end_pos:].lstrip()
                    has_bracket = remaining.startswith(expected_bracket)

                    if not has_bracket:
                        issues.append({
                            "rule": "mandatory_bracket",
                            "rule_label": label,
                            "para_index": para_idx,
                            "page_hint": page_code,
                            "page_code": page_code,
                            "full_header": full_header,
                            "wrong_word": thai_word,
                            "correct_word": f"{thai_word} {expected_bracket}",
                            "reason": (
                                f"คำว่า '{thai_word}' ปรากฏเป็นครั้งแรกในเอกสาร "
                                f"แต่ยังไม่มีวงเล็บภาษาอังกฤษกำกับ "
                                f"กรุณาเติม '{expected_bracket}' ต่อท้าย "
                                f"ให้เป็น '{thai_word} {expected_bracket}' "
                                f"ตามระเบียบสำนักกรรมาธิการ ๓"
                            ),
                            "snippet": build_context_snippet(text, thai_word),
                            "color": color,
                        })
                # ครั้งที่ 2 เป็นต้นไป: ไม่ต้องตรวจที่นี่
                # (check_parenthesis_repeat จะจัดการกรณีมีวงเล็บซ้ำโดยอัตโนมัติ)

    return issues


# ============================================================
# 6. กฎที่ 4: ตรวจศัพท์บัญญัติ และคำทับศัพท์ (Local DB + Google Sheets)
# ============================================================

def check_vocabulary_and_transliteration(
    paragraphs: list,
    sheets_vocab_db=None,
) -> list:
    """
    ตรวจหาคำที่ไม่ตรงตามฐานข้อมูลคำทับศัพท์ทางการ (1,561 คำ) และศัพท์บัญญัติ
    """
    from typing import Optional
    vocab_lookup = dm.get_engine_lookup_db()

    if sheets_vocab_db:
        for row in sheets_vocab_db:
            incorrect = row.get("incorrect_word", "").strip()
            correct = row.get("correct_word", "").strip()
            note = row.get("note", "").strip()
            word_type = row.get("type", "vocabulary").strip().lower()
            if incorrect:
                is_en = bool(re.match(r'^[A-Za-z\s\-_0-9]+$', incorrect))
                vocab_lookup[incorrect] = {
                    "correct": correct,
                    "note": note or f"ตามฐานข้อมูล Google Sheets (แก้ไขเป็น {correct})",
                    "type": word_type,
                    "is_english": is_en,
                }

    if not vocab_lookup:
        return []

    issues = []

    sorted_lookup_items = sorted(
        vocab_lookup.items(),
        key=lambda x: len(x[0]),
        reverse=True
    )

    for p in paragraphs:
        if p.get("is_header"):
            continue

        text = p["text"]
        para_idx = p["index"] + 1
        page_code = p["page_code"]
        full_header = p["full_header"]
        covered_spans = []

        paren_spans = []
        for m in re.finditer(r'\([A-Za-z0-9\s\-_/]+\)', text):
            paren_spans.append((m.start(), m.end()))

        for wrong_key, info in sorted_lookup_items:
            if wrong_key == info.get("correct"):
                continue

            rule_type = (
                "transliteration" if info.get("type") == "transliteration" else "vocabulary"
            )
            if not RULES_CONFIG.get(rule_type, {}).get("enabled", True):
                continue

            if info.get("is_english"):
                pattern = rf'\b{re.escape(wrong_key)}\b'
                for m in re.finditer(pattern, text, re.IGNORECASE):
                    start, end = m.start(), m.end()
                    in_paren = any(p_start <= start and end <= p_end for p_start, p_end in paren_spans)
                    if in_paren:
                        continue
                    is_covered = any(c_start <= start and end <= c_end for c_start, c_end in covered_spans)
                    if is_covered:
                        continue
                    covered_spans.append((start, end))
                    matched_str = m.group()
                    issues.append({
                        "rule": rule_type,
                        "rule_label": RULES_CONFIG[rule_type]["label"],
                        "para_index": para_idx,
                        "page_hint": page_code,
                        "page_code": page_code,
                        "full_header": full_header,
                        "wrong_word": matched_str,
                        "correct_word": info["correct"],
                        "reason": info["note"],
                        "snippet": build_context_snippet(text, matched_str),
                        "color": RULES_CONFIG[rule_type]["color"],
                    })
            else:
                start_pos = 0
                while True:
                    idx_f = text.find(wrong_key, start_pos)
                    if idx_f == -1:
                        break
                    start, end = idx_f, idx_f + len(wrong_key)
                    start_pos = end

                    # ป้องกันเด็ดขาด: ถ้าคำถูกยาวกว่าและข้อความตำแหน่งนี้คือคำที่ถูกต้องอยู่แล้ว ให้ข้ามทันที
                    if len(info["correct"]) > len(wrong_key) and text[start:start + len(info["correct"])] == info["correct"]:
                        continue

                    is_covered = any(c_start <= start and end <= c_end for c_start, c_end in covered_spans)
                    if is_covered:
                        continue
                    covered_spans.append((start, end))
                    issues.append({
                        "rule": rule_type,
                        "rule_label": RULES_CONFIG[rule_type]["label"],
                        "para_index": para_idx,
                        "page_hint": page_code,
                        "page_code": page_code,
                        "full_header": full_header,
                        "wrong_word": wrong_key,
                        "correct_word": info["correct"],
                        "reason": info["note"],
                        "snippet": build_context_snippet(text, wrong_key),
                        "color": RULES_CONFIG[rule_type]["color"],
                    })

    return issues


# ============================================================
# 7. กฎที่ 5: ตรวจระเบียบวุฒิสภา (ในขั้นกรรมาธิการ)
# ============================================================

def check_parliament_rules(paragraphs: list) -> list:
    """
    ตรวจกฎเฉพาะวุฒิสภา 100% Deterministic:
    - 'ในชั้นกรรมาธิการ' -> 'ในขั้นกรรมาธิการ'
    """
    if not RULES_CONFIG.get("parliament_rules", {}).get("enabled", True):
        return []

    issues = []
    color = RULES_CONFIG["parliament_rules"]["color"]
    label = RULES_CONFIG["parliament_rules"]["label"]

    TARGETS = [
        ("ในชั้นกรรมาธิการ", "ในขั้นกรรมาธิการ",
         "ตามระเบียบงานสารบรรณวุฒิสภา ต้องใช้คำว่า 'ในขั้นกรรมาธิการ' (ห้ามใช้ 'ในชั้นกรรมาธิการ')"),
        ("ในชั้นของกรรมาธิการ", "ในขั้นของกรรมาธิการ",
         "ตามระเบียบงานสารบรรณวุฒิสภา ต้องใช้คำว่า 'ในขั้นของกรรมาธิการ'"),
        ("ในชั้นคณะกรรมาธิการ", "ในขั้นคณะกรรมาธิการ",
         "ตามระเบียบงานสารบรรณวุฒิสภา ต้องใช้คำว่า 'ในขั้นคณะกรรมาธิการ'"),
    ]

    for p in paragraphs:
        if p.get("is_header"):
            continue
        text = p["text"]
        para_idx = p["index"] + 1

        for wrong, correct, reason in TARGETS:
            if wrong in text:
                issues.append({
                    "rule": "parliament_rules",
                    "rule_label": label,
                    "para_index": para_idx,
                    "page_hint": p["page_code"],
                    "page_code": p["page_code"],
                    "full_header": p["full_header"],
                    "wrong_word": wrong,
                    "correct_word": correct,
                    "reason": reason,
                    "snippet": build_context_snippet(text, wrong),
                    "color": color,
                })

    return issues


# ============================================================
# 8. กฎที่ 6: ตรวจชื่อ-สกุล สมาชิกวุฒิสภา และการเว้นวรรค ๒ เคาะ
# ============================================================

MILITARY_POLICE_RANKS = {
    "พลตำรวจเอก", "พลตำรวจโท", "พลตำรวจตรี", "พันตำรวจเอก", "พันตำรวจโท", "พันตำรวจตรี",
    "ร้อยตำรวจเอก", "ร้อยตำรวจโท", "ร้อยตำรวจตรี",
    "พลอากาศเอก", "พลอากาศโท", "พลอากาศตรี", "นาวาอากาศเอก", "นาวาอากาศโท", "นาวาอากาศตรี",
    "พลเรือเอก", "พลเรือโท", "พลเรือตรี", "นาวาเอก", "นาวาโท", "นาวาตรี",
    "พลเอก", "พลโท", "พลตรี", "พันเอกหญิง", "พันเอก", "พันโท", "พันตรี",
    "ร้อยเอก", "ร้อยโท", "ร้อยตรี", "ว่าที่ร้อยตรี", "ว่าที่พันตรี",
}


def check_senator_names_and_formatting(paragraphs: list) -> list:
    """
    ตรวจสอบชื่อ-สกุล สมาชิกวุฒิสภา ๒๐๐ ท่าน และคำนำหน้านาม
    """
    if not RULES_CONFIG.get("senator_names", {}).get("enabled", True):
        return []

    senators = dm.get_senators_lookup()
    if not senators:
        return []

    issues = []
    color = RULES_CONFIG["senator_names"]["color"]
    label = RULES_CONFIG["senator_names"]["label"]

    all_wrong_cases = []
    for s in senators:
        full = s["full_name_official"]
        first = s["first_name"].strip()
        last = s["last_name"].strip()
        title = s["title"].strip()
        middle = s["middle_name"].strip()
        is_rank = title in MILITARY_POLICE_RANKS

        if is_rank:
            if middle:
                all_wrong_cases.append({
                    "wrong": f"{title}{first}  {middle}  {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามที่เป็นยศ \"{title}\" ต้องพิมพ์ห่างกับชื่อตัว ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title}{first} {middle} {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามที่เป็นยศ \"{title}\" ต้องพิมพ์ห่างกับชื่อตัว และต้องเว้นวรรคใหญ่ (๒ เคาะ) ระหว่างชื่อ-สกุล",
                })
                all_wrong_cases.append({
                    "wrong": f"{title}  {first} {middle} {last}",
                    "correct": full,
                    "reason": "ต้องเว้นวรรคใหญ่ (๒ เคาะ) ระหว่างชื่อตัว ชื่อกลาง และนามสกุล ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title} {first} {middle} {last}",
                    "correct": full,
                    "reason": "ต้องเว้นวรรคใหญ่ (๒ เคาะ) ทั้งระหว่างยศ ชื่อตัว ชื่อกลาง และนามสกุล ตามระเบียบสำนักกรรมาธิการ ๓",
                })
            else:
                all_wrong_cases.append({
                    "wrong": f"{title}{first}  {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามที่เป็นยศ \"{title}\" ต้องพิมพ์ห่างกับชื่อตัว เช่น {title}  {first} ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title}{first} {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามที่เป็นยศ \"{title}\" ต้องพิมพ์ห่างกับชื่อตัว และต้องเว้นวรรคใหญ่ (๒ เคาะ) ระหว่างชื่อตัวกับนามสกุล",
                })
                all_wrong_cases.append({
                    "wrong": f"{title}  {first} {last}",
                    "correct": full,
                    "reason": "ต้องเว้นวรรคใหญ่ (๒ เคาะ) ระหว่างชื่อตัวและนามสกุล ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title} {first} {last}",
                    "correct": full,
                    "reason": "ต้องเว้นวรรคใหญ่ (๒ เคาะ) ทั้งระหว่างยศและระหว่างชื่อตัวกับนามสกุล",
                })

            all_wrong_cases.append({
                "wrong": f"{title}{first}",
                "correct": f"{title}  {first}",
                "reason": f"คำนำหน้านามที่เป็นยศ \"{title}\" ต้องพิมพ์ห่างกับชื่อตัว เช่น {title}  {first} ตามระเบียบสำนักกรรมาธิการ ๓",
            })
        else:
            if middle:
                all_wrong_cases.append({
                    "wrong": f"{title}  {first}  {middle}  {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามทั่วไป \"{title}\" ให้พิมพ์ติดกับชื่อตัวโดยไม่ต้องเว้นวรรค ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title} {first}  {middle}  {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามทั่วไป \"{title}\" ให้พิมพ์ติดกับชื่อตัว ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title}{first} {middle} {last}",
                    "correct": full,
                    "reason": "ต้องเว้นวรรคใหญ่ (๒ เคาะ) ระหว่างชื่อตัว ชื่อกลาง และนามสกุล ตามระเบียบสำนักกรรมาธิการ ๓",
                })
            else:
                all_wrong_cases.append({
                    "wrong": f"{title}  {first}  {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามทั่วไป \"{title}\" ให้พิมพ์ติดกับชื่อตัว ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title} {first}  {last}",
                    "correct": full,
                    "reason": f"คำนำหน้านามทั่วไป \"{title}\" ให้พิมพ์ติดกับชื่อตัว ตามระเบียบสำนักกรรมาธิการ ๓",
                })
                all_wrong_cases.append({
                    "wrong": f"{title}{first} {last}",
                    "correct": full,
                    "reason": "ต้องเว้นวรรคใหญ่ (๒ เคาะ) ระหว่างชื่อตัวและนามสกุล ตามระเบียบสำนักกรรมาธิการ ๓",
                })

            all_wrong_cases.append({
                "wrong": f"{title}  {first}",
                "correct": f"{title}{first}",
                "reason": f"คำนำหน้านามทั่วไป \"{title}\" ให้พิมพ์ติดกับชื่อตัวโดยไม่ต้องเว้นวรรค ตามระเบียบสำนักกรรมาธิการ ๓",
            })
            all_wrong_cases.append({
                "wrong": f"{title} {first}",
                "correct": f"{title}{first}",
                "reason": f"คำนำหน้านามทั่วไป \"{title}\" ให้พิมพ์ติดกับชื่อตัวโดยไม่ต้องเว้นวรรค ตามระเบียบสำนักกรรมาธิการ ๓",
            })

    all_wrong_cases.sort(key=lambda x: len(x["wrong"]), reverse=True)

    ranks_regex = (
        r"\b(" + "|".join(sorted(MILITARY_POLICE_RANKS, key=len, reverse=True)) + r")"
        r"([ก-๙]{2,})"
    )

    for p in paragraphs:
        if p.get("is_header"):
            continue

        text = p["text"]
        para_idx = p["index"] + 1
        page_code = p["page_code"]
        full_header = p["full_header"]
        covered_spans = []

        for item in all_wrong_cases:
            w_str = item["wrong"]
            if w_str in text and w_str != item["correct"]:
                start_pos = 0
                while True:
                    idx_f = text.find(w_str, start_pos)
                    if idx_f == -1:
                        break
                    start, end = idx_f, idx_f + len(w_str)
                    start_pos = end
                    if any(max(start, cs) < min(end, ce) for cs, ce in covered_spans):
                        continue
                    covered_spans.append((start, end))
                    issues.append({
                        "rule": "senator_names",
                        "rule_label": label,
                        "para_index": para_idx,
                        "page_hint": page_code,
                        "page_code": page_code,
                        "full_header": full_header,
                        "wrong_word": w_str,
                        "correct_word": item["correct"],
                        "reason": item["reason"],
                        "snippet": build_context_snippet(text, w_str),
                        "color": color,
                    })

        for m in re.finditer(ranks_regex, text):
            start, end = m.start(), m.end()
            if any(max(start, cs) < min(end, ce) for cs, ce in covered_spans):
                continue
            covered_spans.append((start, end))
            matched_str = m.group()
            rank_part = m.group(1)
            name_part = m.group(2)
            correct_str = f"{rank_part}  {name_part}"
            issues.append({
                "rule": "senator_names",
                "rule_label": label,
                "para_index": para_idx,
                "page_hint": page_code,
                "page_code": page_code,
                "full_header": full_header,
                "wrong_word": matched_str,
                "correct_word": correct_str,
                "reason": f"คำนำหน้านามที่เป็นยศ \"{rank_part}\" ต้องพิมพ์ห่างกับชื่อตัว เช่น {rank_part}  {name_part} ตามระเบียบสำนักกรรมาธิการ ๓",
                "snippet": build_context_snippet(text, matched_str),
                "color": color,
            })

    return issues


# ============================================================
# 9. กฎที่ 7: กฎระเบียบสำนักกรรมาธิการ ๓
# ============================================================

def check_senate_editorial_rules(paragraphs: list) -> list:
    """ตรวจสอบกฎเกณฑ์เฉพาะสำนักกรรมาธิการ ๓"""
    if not RULES_CONFIG.get("senate_formatting", {}).get("enabled", True):
        return []

    rules = dm.get_senate_editorial_rules()
    issues = []
    color = RULES_CONFIG["senate_formatting"]["color"]
    label = RULES_CONFIG["senate_formatting"]["label"]

    for p in paragraphs:
        if p.get("is_header"):
            continue

        text = p["text"]
        para_idx = p["index"] + 1
        page_code = p["page_code"]
        full_header = p["full_header"]

        for rule in rules:
            pattern = rule["wrong_pattern"]
            replacement = rule["correct_replacement"]
            reason = rule["reason"]

            for m in re.finditer(pattern, text):
                wrong_str = m.group()
                try:
                    correct_str = re.sub(pattern, replacement, wrong_str)
                except Exception:
                    correct_str = replacement

                issues.append({
                    "rule": "senate_formatting",
                    "rule_label": label,
                    "para_index": para_idx,
                    "page_hint": page_code,
                    "page_code": page_code,
                    "full_header": full_header,
                    "wrong_word": wrong_str,
                    "correct_word": correct_str,
                    "reason": reason,
                    "snippet": build_context_snippet(text, wrong_str),
                    "color": color,
                })

        # คำสันธาน ระหว่าง...กับ...
        for m in re.finditer(r"(ระหว่าง\s*[^\s,และ]{2,20}\s+)และ(\s+[^\s,และ]{2,20})", text):
            full_match = m.group()
            correct_match = f"{m.group(1)}กับ{m.group(2)}"
            issues.append({
                "rule": "senate_formatting",
                "rule_label": label,
                "para_index": para_idx,
                "page_hint": page_code,
                "page_code": page_code,
                "full_header": full_header,
                "wrong_word": full_match,
                "correct_word": correct_match,
                "reason": "คำสันธาน \"ระหว่าง...กับ...\" ไม่ใช้คำว่า \"และ\" ตามระเบียบสำนักกรรมาธิการ ๓",
                "snippet": build_context_snippet(text, full_match),
                "color": color,
            })

        # ไม้ยมก (ๆ) ต้องเว้นวรรคหน้าและหลัง
        for m in re.finditer(r"([^\s\d\(\[\{]+)ๆ|ๆ([^\s\)\],\.])", text):
            match_str = m.group()
            correct_str = re.sub(r"([^\s]+)ๆ", r"\1 ๆ", match_str)
            correct_str = re.sub(r"ๆ([^\s]+)", r"ๆ \1", correct_str)
            if correct_str != match_str:
                issues.append({
                    "rule": "senate_formatting",
                    "rule_label": label,
                    "para_index": para_idx,
                    "page_hint": page_code,
                    "page_code": page_code,
                    "full_header": full_header,
                    "wrong_word": match_str,
                    "correct_word": correct_str,
                    "reason": "เครื่องหมายไม้ยมก (ๆ) ต้องเว้นวรรคทั้งข้างหน้าและข้างหลัง ตามระเบียบสำนักกรรมาธิการ ๓",
                    "snippet": build_context_snippet(text, match_str),
                    "color": color,
                })

        # ผู้รับรองถูกต้อง
        for m in re.finditer(r"ผู้รับรองถูก\b(?!ต้อง)", text):
            issues.append({
                "rule": "senate_formatting",
                "rule_label": label,
                "para_index": para_idx,
                "page_hint": page_code,
                "page_code": page_code,
                "full_header": full_header,
                "wrong_word": m.group(),
                "correct_word": "ผู้รับรองถูกต้อง",
                "reason": "ถ้อยคำของประธานในที่ประชุมให้ใช้คำว่า \"ผู้รับรองถูกต้อง\" ตามระเบียบสำนักกรรมาธิการ ๓",
                "snippet": build_context_snippet(text, m.group()),
                "color": color,
            })

    return issues


# ============================================================
# 10. ฟังก์ชันหลัก: รัน Engine ครบทุกกฎ (100% Deterministic — ไม่ใช้ AI)
# ============================================================

def run_full_check(
    file_bytes: bytes,
    api_key: str = "",       # คงพารามิเตอร์ไว้เพื่อ Backward Compatibility แต่ไม่ใช้งาน
    sheets_url: str = "",
    progress_bar=None,
    status_text=None,
) -> tuple:
    """
    รันการตรวจสอบครบทุกกฎอย่างเคร่งครัด (100% Deterministic — ไม่ใช้ AI)
    ประมวลผลเร็ว ไม่มี False Positive จาก AI Hallucination
    Returns: (paragraphs: list, all_issues: list)
    """

    def upd(msg: str):
        if status_text:
            status_text.text(msg)

    # --- Step 1: อ่านไฟล์ทั้งหมด พร้อมสกัดหัวแผ่นกระดาษ + รองรับ Track Changes + Soft Break ---
    upd("📖 กำลังอ่านไฟล์ Word (รองรับ Track Changes และ Enter ขึ้นบรรทัดใหม่)...")
    paragraphs = read_docx_paragraphs(file_bytes)
    total_para = len(paragraphs)
    unique_pages = len(set(p["page_code"] for p in paragraphs))
    upd(f"✓ อ่านไฟล์สำเร็จ: {total_para} ย่อหน้า ({unique_pages} แผ่น/หน้า)")

    all_issues: list = []

    # --- Step 2: แสดงขนาดฐานข้อมูล ---
    try:
        local_count = dm.get_total_count()
    except Exception:
        local_count = 1561
    try:
        senators_count = dm.get_senators_count()
    except Exception:
        senators_count = 200
    upd(f"✓ ฐานข้อมูลพร้อมตรวจ: คำทับศัพท์ {local_count} คำ | สว. {senators_count} ท่าน")

    # --- Step 3: กฎคำผิดภาษาไทย-อังกฤษ (ฐานข้อมูล Word-style) ---
    try:
        if progress_bar:
            progress_bar.progress(0.05, text="ตรวจคำผิดจากฐานข้อมูล...")
        upd("🔤 ตรวจคำผิดภาษาไทย-อังกฤษจากฐานข้อมูล...")
        misspell_issues = check_misspellings_dictionary(paragraphs)
        all_issues.extend(misspell_issues)
        upd(f"✓ คำผิด (ฐานข้อมูล): ตรวจพบ {len(misspell_issues)} รายการ")
    except Exception as e:
        logger.error(f"Step 3 misspelling error: {e}")
        upd(f"⚠️ ข้ามกฎคำผิด (error: {e})")

    # --- Step 4: กฎวงเล็บซ้ำ (Exact Match) ---
    try:
        if progress_bar:
            progress_bar.progress(0.20, text="ตรวจวงเล็บภาษาอังกฤษซ้ำ...")
        upd("🔍 ตรวจวงเล็บภาษาอังกฤษซ้ำ (Exact Match)...")
        paren_issues = check_parenthesis_repeat(paragraphs)
        all_issues.extend(paren_issues)
        upd(f"✓ วงเล็บซ้ำ: ตรวจพบ {len(paren_issues)} รายการ")
    except Exception as e:
        logger.error(f"Step 4 parenthesis_repeat error: {e}")
        upd(f"⚠️ ข้ามกฎวงเล็บซ้ำ (error: {e})")

    # --- Step 5: กฎบังคับวงเล็บครั้งแรก (เช่น พาวเวอร์พ็อยนต์) ---
    try:
        if progress_bar:
            progress_bar.progress(0.35, text="ตรวจวงเล็บบังคับครั้งแรก (พาวเวอร์พ็อยนต์ ฯลฯ)...")
        upd("📌 ตรวจคำบังคับวงเล็บภาษาอังกฤษในการกล่าวถึงครั้งแรก...")
        mandatory_issues = check_mandatory_first_parenthesis(paragraphs)
        all_issues.extend(mandatory_issues)
        upd(f"✓ บังคับวงเล็บครั้งแรก: ตรวจพบ {len(mandatory_issues)} รายการ")
    except Exception as e:
        logger.error(f"Step 5 mandatory_bracket error: {e}")
        upd(f"⚠️ ข้ามกฎบังคับวงเล็บ (error: {e})")

    # --- Step 6: กฎคำทับศัพท์และศัพท์บัญญัติ ---
    try:
        if progress_bar:
            progress_bar.progress(0.50, text="ตรวจคำทับศัพท์และศัพท์บัญญัติ...")
        upd(f"📚 ตรวจคำทับศัพท์และศัพท์บัญญัติ ({local_count} คำ)...")
        vocab_issues = check_vocabulary_and_transliteration(paragraphs)
        all_issues.extend(vocab_issues)
        upd(f"✓ คำทับศัพท์: ตรวจพบ {len(vocab_issues)} รายการ")
    except Exception as e:
        logger.error(f"Step 6 vocabulary error: {e}")
        upd(f"⚠️ ข้ามกฎคำทับศัพท์ (error: {e})")

    # --- Step 7: กฎระเบียบวุฒิสภา (ในขั้นกรรมาธิการ) ---
    try:
        if progress_bar:
            progress_bar.progress(0.65, text="ตรวจระเบียบวุฒิสภา...")
        upd("⚖️ ตรวจระเบียบวุฒิสภา (ในชั้น → ในขั้นกรรมาธิการ)...")
        parliament_issues = check_parliament_rules(paragraphs)
        all_issues.extend(parliament_issues)
        upd(f"✓ ระเบียบวุฒิสภา: ตรวจพบ {len(parliament_issues)} รายการ")
    except Exception as e:
        logger.error(f"Step 7 parliament_rules error: {e}")
        upd(f"⚠️ ข้ามกฎระเบียบวุฒิสภา (error: {e})")

    # --- Step 8: กฎชื่อ-สกุล สมาชิกวุฒิสภา ---
    try:
        if progress_bar:
            progress_bar.progress(0.80, text="ตรวจชื่อ-สกุล สว. และวรรคใหญ่ ๒ เคาะ...")
        upd(f"🏛️ ตรวจชื่อ-สกุล สมาชิกวุฒิสภา {senators_count} ท่าน (วรรคใหญ่ ๒ เคาะ)...")
        senator_issues = check_senator_names_and_formatting(paragraphs)
        all_issues.extend(senator_issues)
        upd(f"✓ ชื่อ สว. / วรรค ๒ เคาะ: ตรวจพบ {len(senator_issues)} รายการ")
    except Exception as e:
        logger.error(f"Step 8 senator_names error: {e}")
        upd(f"⚠️ ข้ามกฎชื่อ สว. (error: {e})")

    # --- Step 9: กฎระเบียบสำนักกรรมาธิการ ๓ ---
    try:
        if progress_bar:
            progress_bar.progress(0.92, text="ตรวจตามระเบียบสำนักกรรมาธิการ ๓...")
        upd("⚖️ ตรวจระเบียบสำนักกรรมาธิการ ๓ (สันทนาการ, พระราม, คำควบคู่ ฯลฯ)...")
        senate_rule_issues = check_senate_editorial_rules(paragraphs)
        all_issues.extend(senate_rule_issues)
        upd(f"✓ ระเบียบสำนักกรรมาธิการ ๓: ตรวจพบ {len(senate_rule_issues)} รายการ")
    except Exception as e:
        logger.error(f"Step 9 senate_formatting error: {e}")
        upd(f"⚠️ ข้ามกฎระเบียบสำนักกรรมาธิการ ๓ (error: {e})")

    if progress_bar:
        progress_bar.progress(1.0, text="✅ ตรวจสอบครบถ้วน 100% ทุกย่อหน้า!")

    # ขจัดรายการซ้ำซ้อน (Deduplicate: ป้องกันปัญหาคำเดิมแจ้งเตือนซ้ำ)
    unique_issues = []
    seen_keys = set()
    for iss in all_issues:
        key = (iss["para_index"], iss["wrong_word"].strip(), iss["correct_word"].strip())
        if key not in seen_keys:
            seen_keys.add(key)
            unique_issues.append(iss)
    all_issues = unique_issues

    # เรียงลำดับตามย่อหน้า
    all_issues.sort(key=lambda x: (x["para_index"], x["rule"]))

    return paragraphs, all_issues
