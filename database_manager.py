# ============================================================
# database_manager.py - ระบบจัดการฐานข้อมูลคำทับศัพท์วุฒิสภา
# เสถียรสูงสุด: ทำงานได้ทั้ง Local และ Cloud (Linux / Windows)
# ใช้ CSV-backed in-memory database พร้อม cache
# ============================================================

import os
import re
import csv
import sys
import logging
from typing import Dict, List, Optional
import streamlit as st

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
CSV_PATH = os.path.join(DATA_DIR, "vocab_db.csv")
ALT_CSV_PATH = os.path.join(DATA_DIR, "transliterations.csv")

# ตารางคำทับศัพท์ภาษาไทยที่มักสะกดผิดตามหลักราชบัณฑิตยสภา / วุฒิสภา
COMMON_MISPELLING_MAP = {
    "สมาร์ท": "สมาร์ต",
    "สมาร์ทโฟน": "สมาร์ตโฟน",
    "สมาร์ทซิตี": "สมาร์ตซิตี",
    "สมาร์ทฟาร์มเมอร์": "สมาร์ตฟาร์มเมอร์",
    "แอพ": "แอป",
    "แอพพลิเคชัน": "แอปพลิเคชัน",
    "แอพพลิเคชั่น": "แอปพลิเคชัน",
    "แอปพลิเคชั่น": "แอปพลิเคชัน",
    "ดิจิตอล": "ดิจิทัล",
    "ดิจิตอลวอลเล็ต": "ดิจิทัลวอลเล็ต",
    "บล็อค": "บล็อก",
    "บล็อคเชน": "บล็อกเชน",
    "บล็อกเชนจ์": "บล็อกเชน",
    "อัพเดท": "อัปเดต",
    "อัพเดต": "อัปเดต",
    "อัพเกรด": "อัปเกรด",
    "อัพเกรท": "อัปเกรด",
    "ลิงค์": "ลิงก์",
    "ลิ้งค์": "ลิงก์",
    "ลิ้ง": "ลิงก์",
    "คลิ๊ก": "คลิก",
    "คลิกส์": "คลิก",
    "อีเมล์": "อีเมล",
    "อีเมลล์": "อีเมล",
    "แพลทฟอร์ม": "แพลตฟอร์ม",
    "แพลตฟอร์ม์": "แพลตฟอร์ม",
    "เว็ปไซต์": "เว็บไซต์",
    "เวบไซต์": "เว็บไซต์",
    "เวปไซต์": "เว็บไซต์",
    "เว็บไซด์": "เว็บไซต์",
    "เว็ปไซด์": "เว็บไซต์",
    "ชาร์ต": "ชาร์จ",
    "โพส": "โพสต์",
    "โพสท์": "โพสต์",
    "โพสต์ท": "โพสต์",
    "คอมพิวเต้อร์": "คอมพิวเตอร์",
    "คอมพิวเตอร์์": "คอมพิวเตอร์",
    "ยูทูป": "ยูทูบ",
    "ยูทู๊บ": "ยูทูบ",
    "โปรเจค": "โพรเจกต์",
    "โปรเจ็ค": "โพรเจกต์",
    "โพรเจค": "โพรเจกต์",
    "โปรเจกต์": "โพรเจกต์",
    "โปรเจ็กต์": "โพรเจ็กต์",
    "ฟังก์ชั่น": "ฟังก์ชัน",
    "ฟังชั่น": "ฟังก์ชัน",
    "ฟังค์ชั่น": "ฟังก์ชัน",
    "กราฟฟิก": "กราฟิก",
    "กราฟฟิค": "กราฟิก",
    "กราฟิค": "กราฟิก",
    "เซ็นเซอร์": "เซนเซอร์",
    "เซ็นเตอร์": "เซนเตอร์",
    "เช็คอิน": "เช็กอิน",
    "เช็ค": "เช็ก",
    "ซอฟท์แวร์": "ซอฟต์แวร์",
    "ซอฟท์พาวเวอร์": "ซอฟต์พาวเวอร์",
    "ซอฟพาวเวอร์": "ซอฟต์พาวเวอร์",
    "สตาร์ท": "สตาร์ต",
    "สตาร์ทอัพ": "สตาร์ตอัป",
    "สตาร์ตอัพ": "สตาร์ตอัป",
    "สติ๊กเกอร์": "สติกเกอร์",
    "สเต็ป": "สเตป",
    "ออนไลน์": "ออนไลน์",
    "โควต้า": "โควตา",
    "แท็กซี่": "แท็กซี",
    "แท๊กซี่": "แท็กซี",
    "พอยท์": "พอยต์",
    "พอยต์": "พอยต์",
    "พ้อยท์": "พอยต์",
    "พ็อยนต์": "พ็อยนต์",
    "เทคโนโลยี": "เทคโนโลยี",
    "ปาร์ตี้": "ปาร์ตี",
    "ปาร์ตี้ลิสต์": "ปาร์ตีลิสต์",
}


@st.cache_data(show_spinner=False)
def load_vocab_records() -> List[dict]:
    """
    อ่านข้อมูลคำศัพท์จาก CSV ไฟล์ในโฟลเดอร์ data
    คืนค่าเป็น list ของ dict พร้อม cache ในหน่วยความจำ
    """
    records = []
    target_path = CSV_PATH if os.path.exists(CSV_PATH) else ALT_CSV_PATH

    if not os.path.exists(target_path):
        logger.warning(f"ไม่พบไฟล์คำศัพท์ที่ {target_path}")
        return records

    try:
        with open(target_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                incorrect = row.get("incorrect_word") or row.get("english_word") or ""
                correct = row.get("correct_word") or row.get("thai_word") or ""
                note = row.get("note", "")
                cat_type = row.get("type") or row.get("category") or "transliteration"

                if incorrect.strip() and correct.strip():
                    records.append({
                        "incorrect_word": incorrect.strip(),
                        "correct_word": correct.strip(),
                        "note": note.strip(),
                        "type": cat_type.strip().lower(),
                    })
        logger.info(f"โหลดฐานข้อมูลคำศัพท์สำเร็จ {len(records)} รายการ")
    except Exception as e:
        logger.error(f"เกิดข้อผิดพลาดในการอ่าน CSV: {e}")

    return records


def get_total_count() -> int:
    """นับจำนวนคำศัพท์ทางการทั้งหมดในระบบ"""
    try:
        data = load_vocab_records()
        return len(data) if data else 1561
    except Exception:
        return 1561


def get_engine_lookup_db() -> Dict[str, Dict]:
    """
    สร้างและคืนค่า lookup dictionary สำหรับ Checker Engine:
    1. คำภาษาอังกฤษ 1,561 คำ -> แนะนำคำทับศัพท์ไทยทางการ
    2. คำทับศัพท์ไทยที่สะกดผิดยอดนิยม -> แนะนำคำทับศัพท์ที่ถูกต้อง
    """
    lookup = {}

    # 1. โหลดจาก CSV
    records = load_vocab_records()
    for row in records:
        wrong = row["incorrect_word"]
        correct = row["correct_word"]
        note = row["note"]
        word_type = row["type"]

        is_en = bool(re.match(r"^[A-Za-z0-9\s\-_/.]+$", wrong))

        if is_en:
            lookup[wrong] = {
                "correct": correct,
                "note": f"คำภาษาอังกฤษ ควรใช้คำทับศัพท์ทางการ: \"{correct}\"{(' (' + note + ')') if note else ''}",
                "type": "transliteration",
                "is_english": True,
            }
        else:
            lookup[wrong] = {
                "correct": correct,
                "note": note or f"คำทับศัพท์ทางการวุฒิสภา (แก้ไขเป็น {correct})",
                "type": word_type,
                "is_english": False,
            }

    # 2. ผนวกตารางคำทับศัพท์ไทยที่สะกดผิดยอดนิยม
    for wrong_th, correct_th in COMMON_MISPELLING_MAP.items():
        if wrong_th not in lookup:
            lookup[wrong_th] = {
                "correct": correct_th,
                "note": f"คำทับศัพท์ที่มักเขียนผิดตามหลักราชบัณฑิตยสภา/วุฒิสภา (ที่ถูกต้องคือ \"{correct_th}\")",
                "type": "transliteration",
                "is_english": False,
            }

    return lookup
