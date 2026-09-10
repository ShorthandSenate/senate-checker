# ============================================================
# checker_engine.py - หัวใจของการตรวจสอบ (Engine Layer)
# ============================================================

import re
import io
import time
import json
import logging
import requests
import streamlit as st
import google.generativeai as genai
from docx import Document
from typing import Optional

from config import (
    GEMINI_MODELS,
    RETRY_MAX_ATTEMPTS,
    RETRY_WAIT_MIN_SEC,
    RETRY_WAIT_MAX_SEC,
    BATCH_SIZE,
    CONTEXT_WORDS,
    SHEETS_CACHE_TTL,
    RULES_CONFIG,
    SPELLING_AI_PROMPT,
)

logger = logging.getLogger(__name__)

# ============================================================
# 1. โหลดฐานข้อมูลจาก Google Sheets (with cache)
# ============================================================

@st.cache_data(ttl=SHEETS_CACHE_TTL, show_spinner=False)
def load_vocabulary_db(sheets_csv_url: str) -> list:
    """ดึงข้อมูลคำศัพท์จาก Google Sheets CSV Export URL พร้อม In-memory cache"""
    try:
        resp = requests.get(sheets_csv_url, timeout=15)
        resp.raise_for_status()
        lines = resp.text.strip().splitlines()
        if len(lines) < 2:
            return []
        headers = [h.strip().strip('"') for h in lines[0].split(",")]
        rows = []
        for line in lines[1:]:
            parts = line.split(",", maxsplit=len(headers) - 1)
            if len(parts) >= 2:
                row = {
                    headers[i]: parts[i].strip().strip('"')
                    for i in range(min(len(headers), len(parts)))
                }
                rows.append(row)
        logger.info(f"โหลด Google Sheets สำเร็จ {len(rows)} รายการ")
        return rows
    except Exception as e:
        logger.warning(f"โหลด Google Sheets ไม่สำเร็จ: {e}")
        return []


# ============================================================
# 2. อ่านไฟล์ Word และแยกย่อหน้า (100% ครบทุกย่อหน้า)
# ============================================================

def read_docx_paragraphs(file_bytes: bytes) -> list:
    """
    อ่านไฟล์ .docx และแยกย่อหน้าทุกย่อหน้า (ไม่ข้าม)
    คืนค่า list ของ dict: {index, text, page_hint}
    """
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    para_index = 0
    page_count = 1

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # ตรวจ page break จาก XML
        xml_str = para._element.xml
        if 'w:type="page"' in xml_str:
            page_count += 1

        paragraphs.append({
            "index": para_index,
            "text": text,
            "page_hint": page_count,
        })
        para_index += 1

    logger.info(f"อ่านไฟล์สำเร็จ: {para_index} ย่อหน้า, ~{page_count} หน้า")
    return paragraphs


# ============================================================
# 3. Multi-Model Fallback AI + Exponential Backoff
# ============================================================

def call_gemini_with_fallback(prompt: str, api_key: str) -> Optional[str]:
    """
    ส่ง prompt ไปยัง Gemini พร้อมระบบ:
    - Multi-Model Fallback (3 โมเดล)
    - Exponential Backoff (2-15 วินาที)
    - Auto-retry per model (3 ครั้ง)
    คืนค่า text response หรือ None ถ้าล้มเหลวทุกโมเดล
    """
    genai.configure(api_key=api_key)
    RETRIABLE_CODES = {"429", "500", "503"}

    for model_name in GEMINI_MODELS:
        for attempt in range(1, RETRY_MAX_ATTEMPTS + 1):
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content(
                    prompt,
                    generation_config=genai.types.GenerationConfig(
                        temperature=0.1,
                        max_output_tokens=2048,
                    ),
                )
                if response and response.text:
                    logger.info(f"✓ AI สำเร็จ: {model_name} (attempt {attempt})")
                    return response.text
            except Exception as e:
                err_str = str(e)
                is_retriable = any(code in err_str for code in RETRIABLE_CODES)

                if is_retriable and attempt < RETRY_MAX_ATTEMPTS:
                    wait_sec = min(
                        RETRY_WAIT_MIN_SEC * (2 ** (attempt - 1)),
                        RETRY_WAIT_MAX_SEC,
                    )
                    logger.warning(
                        f"⚠ {model_name} attempt {attempt} error ({err_str[:60]}) → หน่วง {wait_sec}s"
                    )
                    time.sleep(wait_sec)
                else:
                    logger.warning(f"✗ {model_name} ล้มเหลว: {err_str[:100]}")
                    break

        logger.warning(f"→ สลับจาก {model_name} ไปโมเดลถัดไป...")

    logger.error("AI ล้มเหลวทุกโมเดล - batch นี้จะถูกข้าม")
    return None


# ============================================================
# 4. Parse JSON จาก AI Response (Robust)
# ============================================================

def parse_ai_json(raw_text: str) -> list:
    """แยก JSON จาก AI response ที่อาจมี markdown code block ปน"""
    if not raw_text:
        return []
    cleaned = re.sub(r"```(?:json)?\s*", "", raw_text).replace("```", "").strip()
    try:
        data = json.loads(cleaned)
        return data.get("issues", [])
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if match:
            try:
                data = json.loads(match.group())
                return data.get("issues", [])
            except Exception:
                pass
    return []


# ============================================================
# 5. Context Snippet Builder
# ============================================================

def build_context_snippet(text: str, wrong_word: str, word_count: int = CONTEXT_WORDS) -> str:
    """ดึงข้อความรอบข้างคำผิด word_count คำทั้งสองข้าง"""
    idx = text.find(wrong_word)
    if idx == -1:
        return (text[:100] + "...") if len(text) > 100 else text

    left_words = text[:idx].split()
    right_words = text[idx + len(wrong_word):].split()
    left_snippet = " ".join(left_words[-word_count:])
    right_snippet = " ".join(right_words[:word_count])
    return f"{left_snippet} [{wrong_word}] {right_snippet}".strip()


# ============================================================
# 6. กฎที่ 1: ตรวจคำผิดทั่วไปด้วย AI (Batch Processing)
# ============================================================

def check_spelling_ai(
    paragraphs: list,
    api_key: str,
    progress_callback=None,
) -> list:
    """ตรวจคำผิดทั่วไปด้วย Gemini AI แบบ batch processing"""
    if not RULES_CONFIG["spelling"]["enabled"] or not api_key:
        return []

    issues = []
    total_batches = (len(paragraphs) + BATCH_SIZE - 1) // BATCH_SIZE

    for batch_num in range(total_batches):
        batch = paragraphs[batch_num * BATCH_SIZE : (batch_num + 1) * BATCH_SIZE]
        batch_text = "\n\n".join(
            f"[ย่อหน้า {p['index'] + 1}]: {p['text']}" for p in batch
        )
        prompt = SPELLING_AI_PROMPT.format(text=batch_text)
        raw_response = call_gemini_with_fallback(prompt, api_key)

        if raw_response:
            for issue in parse_ai_json(raw_response):
                wrong = issue.get("wrong_word", "").strip()
                correct = issue.get("correct_word", "").strip()
                reason = issue.get("reason", "")
                if not wrong or not correct or wrong == correct:
                    continue
                for p in batch:
                    if wrong in p["text"]:
                        issues.append({
                            "rule": "spelling",
                            "rule_label": RULES_CONFIG["spelling"]["label"],
                            "para_index": p["index"] + 1,
                            "page_hint": p["page_hint"],
                            "wrong_word": wrong,
                            "correct_word": correct,
                            "reason": reason,
                            "snippet": build_context_snippet(p["text"], wrong),
                            "color": RULES_CONFIG["spelling"]["color"],
                        })

        if progress_callback:
            progress_callback(batch_num + 1, total_batches)

        # rate limiting buffer
        time.sleep(0.4)

    return issues


# ============================================================
# 7. กฎที่ 2 + 4: ตรวจศัพท์บัญญัติ และคำทับศัพท์ (Local DB + Google Sheets)
# ============================================================

import database_manager as dm

def check_vocabulary_and_transliteration(
    paragraphs: list,
    sheets_vocab_db: Optional[list] = None,
) -> list:
    """
    ตรวจหาคำที่ไม่ตรงตามฐานข้อมูลคำทับศัพท์ทางการ (1,561 คำ) และศัพท์บัญญัติ:
    1. ตรวจจับคำภาษาอังกฤษที่ปรากฏในข้อความ เพื่อแนะนำคำทับศัพท์ไทยทางการ
    2. ตรวจจับคำทับศัพท์ไทยที่สะกดผิดตามหลักราชบัณฑิตยสภา/วุฒิสภา
    3. ผสานข้อมูลจาก Google Sheets (ถ้ามี)
    """
    # 1. โหลดฐานข้อมูลหลักจาก Local Database
    vocab_lookup = dm.get_engine_lookup_db()

    # 2. ผสานจาก Google Sheets ถ้ามี
    if sheets_vocab_db:
        for row in sheets_vocab_db:
            incorrect = row.get("incorrect_word", "").strip()
            correct = row.get("correct_word", "").strip()
            note = row.get("note", "").strip()
            word_type = row.get("type", "vocabulary").strip().lower()
            if incorrect:
                is_en = bool(re.match(r"^[A-Za-z\s\-_0-9]+$", incorrect))
                vocab_lookup[incorrect] = {
                    "correct": correct,
                    "note": note or f"ตามฐานข้อมูล Google Sheets (แก้ไขเป็น {correct})",
                    "type": word_type,
                    "is_english": is_en,
                }

    if not vocab_lookup:
        return []

    issues = []

    # จัดเรียงคำค้นหาจากยาวไปสั้น เพื่อให้จับคำประสมที่ยาวกว่าก่อนเสมอ (Longest Match First)
    # เช่น "soft power" ก่อน "power", "สมาร์ทโฟน" ก่อน "สมาร์ท"
    sorted_lookup_items = sorted(
        vocab_lookup.items(),
        key=lambda x: len(x[0]),
        reverse=True
    )

    for p in paragraphs:
        text = p["text"]
        para_idx = p["index"] + 1

        # เก็บช่วงตัวอักษร (start, end) ที่ถูกตรวจพบไปแล้ว เพื่อไม่ให้ตรวจซ้ำในคำย่อย
        covered_spans = []

        # หาช่วงคำในวงเล็บภาษาอังกฤษ (เช่น เพื่อขยายความคำไทยอย่างถูกต้อง)
        paren_spans = []
        for m in re.finditer(r"\([A-Za-z0-9\s\-_/]+\)", text):
            paren_spans.append((m.start(), m.end()))

        for wrong_key, info in sorted_lookup_items:
            rule_type = (
                "transliteration" if info.get("type") == "transliteration" else "vocabulary"
            )
            if not RULES_CONFIG.get(rule_type, {}).get("enabled", True):
                continue

            if info.get("is_english"):
                # ตรวจคำภาษาอังกฤษแบบเต็มคำ (Word Boundary)
                pattern = rf"\b{re.escape(wrong_key)}\b"
                for m in re.finditer(pattern, text, re.IGNORECASE):
                    start, end = m.start(), m.end()

                    # ข้ามหากอยู่ในวงเล็บขยายความคำแปล
                    in_paren = any(p_start <= start and end <= p_end for p_start, p_end in paren_spans)
                    if in_paren:
                        continue

                    # ข้ามหากช่วงนี้ถูกครอบคลุมโดยคำประสมที่ยาวกว่าแล้ว
                    is_covered = any(c_start <= start and end <= c_end for c_start, c_end in covered_spans)
                    if is_covered:
                        continue

                    covered_spans.append((start, end))
                    matched_str = m.group()

                    issues.append({
                        "rule": rule_type,
                        "rule_label": RULES_CONFIG[rule_type]["label"],
                        "para_index": para_idx,
                        "page_hint": p["page_hint"],
                        "wrong_word": matched_str,
                        "correct_word": info["correct"],
                        "reason": info["note"],
                        "snippet": build_context_snippet(text, matched_str),
                        "color": RULES_CONFIG[rule_type]["color"],
                    })
            else:
                # คำภาษาไทย (เช่น คำทับศัพท์ที่สะกดผิด)
                start_pos = 0
                while True:
                    idx = text.find(wrong_key, start_pos)
                    if idx == -1:
                        break
                    start, end = idx, idx + len(wrong_key)
                    start_pos = end

                    # ข้ามหากช่วงนี้ถูกครอบคลุมโดยคำที่ยาวกว่าแล้ว
                    is_covered = any(c_start <= start and end <= c_end for c_start, c_end in covered_spans)
                    if is_covered:
                        continue

                    covered_spans.append((start, end))

                    issues.append({
                        "rule": rule_type,
                        "rule_label": RULES_CONFIG[rule_type]["label"],
                        "para_index": para_idx,
                        "page_hint": p["page_hint"],
                        "wrong_word": wrong_key,
                        "correct_word": info["correct"],
                        "reason": info["note"],
                        "snippet": build_context_snippet(text, wrong_key),
                        "color": RULES_CONFIG[rule_type]["color"],
                    })

    return issues


# ============================================================
# 8. กฎที่ 3: ตรวจวงเล็บภาษาอังกฤษซ้ำ
# ============================================================

def check_parenthesis_repeat(paragraphs: list) -> list:
    """
    ตรวจจับคำภาษาอังกฤษในวงเล็บที่ปรากฏซ้ำ
    อนุญาตให้ใส่วงเล็บได้ครั้งแรกเท่านั้น
    """
    if not RULES_CONFIG["parenthesis_repeat"]["enabled"]:
        return []

    pattern = RULES_CONFIG["parenthesis_repeat"]["bracket_pattern"]
    seen: dict = {}  # normalized_key -> {para_index, page_hint, display}
    issues = []

    for p in paragraphs:
        text = p["text"]
        matches = re.findall(pattern, text)
        for match in matches:
            key = match.strip().lower()
            display = match.strip()
            bracket_form = f"({display})"

            if key not in seen:
                seen[key] = {
                    "para_index": p["index"] + 1,
                    "page_hint": p["page_hint"],
                    "display": display,
                }
            else:
                first = seen[key]
                issues.append({
                    "rule": "parenthesis_repeat",
                    "rule_label": RULES_CONFIG["parenthesis_repeat"]["label"],
                    "para_index": p["index"] + 1,
                    "page_hint": p["page_hint"],
                    "wrong_word": bracket_form,
                    "correct_word": (
                        f"ตัดวงเล็บออก (ใส่แล้วที่ย่อหน้า {first['para_index']} "
                        f"~หน้า {first['page_hint']})"
                    ),
                    "reason": (
                        f"({display}) ปรากฏครั้งแรกที่ย่อหน้า {first['para_index']} "
                        f"(~หน้า {first['page_hint']}) ไม่ควรใส่วงเล็บซ้ำ"
                    ),
                    "snippet": build_context_snippet(text, bracket_form),
                    "color": RULES_CONFIG["parenthesis_repeat"]["color"],
                })

    return issues


# ============================================================
# 9. ฟังก์ชันหลัก: รัน Engine ครบทุกกฎ
# ============================================================

def run_full_check(
    file_bytes: bytes,
    api_key: str,
    sheets_url: str,
    progress_bar=None,
    status_text=None,
) -> tuple:
    """
    รันการตรวจสอบครบทุกกฎ (100% ทุกย่อหน้า)
    Returns: (paragraphs: list, all_issues: list)
    """

    def upd(msg: str):
        if status_text:
            status_text.text(msg)

    # --- Step 1: อ่านไฟล์ ---
    upd("📖 กำลังอ่านไฟล์ Word...")
    paragraphs = read_docx_paragraphs(file_bytes)
    total_para = len(paragraphs)
    upd(f"✓ อ่านไฟล์สำเร็จ พบ {total_para} ย่อหน้า")

    all_issues: list = []

    # --- Step 2: เตรียมฐานข้อมูลคำทับศัพท์ (Local DB + Google Sheets) ---
    local_count = dm.get_total_count()
    sheets_vocab_db: list = []
    if sheets_url and sheets_url.startswith("http"):
        upd("📊 โหลดฐานข้อมูลเสริมจาก Google Sheets...")
        sheets_vocab_db = load_vocabulary_db(sheets_url)
        upd(f"✓ โหลดฐานข้อมูลทางการ {local_count} คำ + Sheets {len(sheets_vocab_db)} คำ")
    else:
        upd(f"✓ โหลดฐานข้อมูลคำทับศัพท์ทางการ {local_count} คำ (พร้อมตรวจ)")

    # --- Step 3: กฎวงเล็บซ้ำ (เร็ว ไม่ต้องใช้ AI) ---
    if progress_bar:
        progress_bar.progress(0.05, text="ตรวจวงเล็บซ้ำ...")
    upd("🔍 ตรวจวงเล็บซ้ำ...")
    paren_issues = check_parenthesis_repeat(paragraphs)
    all_issues.extend(paren_issues)
    upd(f"✓ วงเล็บซ้ำ: พบ {len(paren_issues)} รายการ")

    # --- Step 4: กฎคำทับศัพท์และศัพท์บัญญัติ (เทียบกับฐานข้อมูลทางการ) ---
    if progress_bar:
        progress_bar.progress(0.15, text="ตรวจศัพท์บัญญัติและคำทับศัพท์...")
    upd("📚 ตรวจศัพท์บัญญัติและคำทับศัพท์เทียบฐานข้อมูล...")
    vocab_issues = check_vocabulary_and_transliteration(paragraphs, sheets_vocab_db)
    all_issues.extend(vocab_issues)
    upd(f"✓ คำทับศัพท์/ศัพท์บัญญัติ: ตรวจพบ {len(vocab_issues)} รายการ")

    # --- Step 5: กฎคำผิดทั่วไปด้วย AI (Batch + Fallback) ---
    if api_key and RULES_CONFIG["spelling"]["enabled"]:
        total_batches = (total_para + BATCH_SIZE - 1) // BATCH_SIZE
        upd(f"🤖 ตรวจคำผิดด้วย AI ({total_batches} batches)...")

        def ai_progress(done, total):
            pct = 0.2 + (done / total) * 0.75
            if progress_bar:
                progress_bar.progress(
                    min(pct, 0.95), text=f"🤖 AI: batch {done}/{total}"
                )
            upd(f"🤖 AI ตรวจ batch {done}/{total} เสร็จแล้ว")

        spelling_issues = check_spelling_ai(paragraphs, api_key, ai_progress)
        all_issues.extend(spelling_issues)
        upd(f"✓ คำผิด AI: พบ {len(spelling_issues)} รายการ")
    else:
        if not api_key:
            upd("⚠️ ข้ามการตรวจ AI (ไม่มี API Key)")

    if progress_bar:
        progress_bar.progress(1.0, text="✅ ตรวจสอบครบ 100%!")

    # เรียงตามลำดับย่อหน้า
    all_issues.sort(key=lambda x: (x["para_index"], x["rule"]))

    return paragraphs, all_issues
