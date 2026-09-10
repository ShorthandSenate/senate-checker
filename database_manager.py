# ============================================================
# database_manager.py - ระบบจัดการฐานข้อมูลคำทับศัพท์วุฒิสภา (CRUD & Flexible Management)
# ============================================================

import os
import sys
import re
import csv
import io
import json
import sqlite3
import logging
from typing import Optional, List, Dict, Tuple
from datetime import datetime

# บังคับใช้ UTF-8 สำหรับ stdout
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
SQLITE_DB_PATH = os.path.join(DATA_DIR, "vocab.db")
SEED_CSV_PATH = os.path.join(DATA_DIR, "transliterations.csv")


def get_connection() -> sqlite3.Connection:
    """เปิด connection ไปยัง SQLite DB พร้อม row_factory"""
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(SQLITE_DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    return conn


def init_db_if_needed():
    """ตรวจสอบและสร้างฐานข้อมูลหากยังไม่มี"""
    if not os.path.exists(SQLITE_DB_PATH) or os.path.getsize(SQLITE_DB_PATH) == 0:
        logger.info("ไม่พบ vocab.db กำลังสร้างจาก database_builder...")
        try:
            from database_builder import main as build_main
            build_main()
        except Exception as e:
            logger.error(f"สร้างฐานข้อมูลล้มเหลว: {e}")


# เรียก init เมื่อโหลดโมดูล
init_db_if_needed()


# ============================================================
# 1. ฟังก์ชันดึงและนับข้อมูล (Read & Query)
# ============================================================

def get_total_count() -> int:
    """นับจำนวนคำศัพท์ที่เปิดใช้งานในฐานข้อมูล"""
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM transliterations WHERE is_active = 1")
        count = cur.fetchone()[0]
        conn.close()
        return count
    except Exception as e:
        logger.error(f"get_total_count error: {e}")
        return 0


def get_all_records(
    search_query: str = "",
    limit: int = 100,
    offset: int = 0
) -> Tuple[List[Dict], int]:
    """
    ดึงรายการคำศัพท์ทั้งหมด รองรับการค้นหา (ทั้ง EN และ TH) และทำ Pagination
    คืนค่า (records, total_matching_count)
    """
    try:
        conn = get_connection()
        cur = conn.cursor()

        if search_query.strip():
            q = f"%{search_query.strip()}%"
            count_sql = """
                SELECT COUNT(*) FROM transliterations 
                WHERE is_active = 1 AND (english_word LIKE ? OR thai_word LIKE ? OR note LIKE ?)
            """
            cur.execute(count_sql, (q, q, q))
            total = cur.fetchone()[0]

            query_sql = """
                SELECT id, english_word, thai_word, note, category, updated_at
                FROM transliterations
                WHERE is_active = 1 AND (english_word LIKE ? OR thai_word LIKE ? OR note LIKE ?)
                ORDER BY id ASC
                LIMIT ? OFFSET ?
            """
            cur.execute(query_sql, (q, q, q, limit, offset))
        else:
            cur.execute("SELECT COUNT(*) FROM transliterations WHERE is_active = 1")
            total = cur.fetchone()[0]

            cur.execute("""
                SELECT id, english_word, thai_word, note, category, updated_at
                FROM transliterations
                WHERE is_active = 1
                ORDER BY id ASC
                LIMIT ? OFFSET ?
            """, (limit, offset))

        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows, total
    except Exception as e:
        logger.error(f"get_all_records error: {e}")
        return [], 0


def get_record_by_id(record_id: int) -> Optional[Dict]:
    """ดึงข้อมูลตาม ID"""
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM transliterations WHERE id = ?", (record_id,))
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception as e:
        logger.error(f"get_record_by_id error: {e}")
        return None


def get_record_by_en(en_word: str) -> Optional[Dict]:
    """ดึงข้อมูลตามคำภาษาอังกฤษ"""
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM transliterations WHERE english_word = ? COLLATE NOCASE AND is_active = 1",
            (en_word.strip(),)
        )
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception as e:
        logger.error(f"get_record_by_en error: {e}")
        return None


# ============================================================
# 2. ฟังก์ชันแก้ไข/เพิ่ม/ลบ (Update, Add, Delete)
# ============================================================

def update_word(
    identifier,  # id: int หรือ english_word: str
    new_thai_word: str,
    new_note: Optional[str] = None
) -> bool:
    """
    แก้ไขเฉพาะบางคำ:
    - เปลี่ยนคำทับศัพท์ภาษาไทยที่ถูกต้อง
    - อัปเดตหมายเหตุ
    """
    try:
        conn = get_connection()
        cur = conn.cursor()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if isinstance(identifier, int) or (isinstance(identifier, str) and identifier.isdigit()):
            rec_id = int(identifier)
            if new_note is not None:
                cur.execute("""
                    UPDATE transliterations 
                    SET thai_word = ?, note = ?, updated_at = ?
                    WHERE id = ?
                """, (new_thai_word.strip(), new_note.strip(), now, rec_id))
            else:
                cur.execute("""
                    UPDATE transliterations 
                    SET thai_word = ?, updated_at = ?
                    WHERE id = ?
                """, (new_thai_word.strip(), now, rec_id))
        else:
            en = str(identifier).strip()
            if new_note is not None:
                cur.execute("""
                    UPDATE transliterations 
                    SET thai_word = ?, note = ?, updated_at = ?
                    WHERE english_word = ? COLLATE NOCASE
                """, (new_thai_word.strip(), new_note.strip(), now, en))
            else:
                cur.execute("""
                    UPDATE transliterations 
                    SET thai_word = ?, updated_at = ?
                    WHERE english_word = ? COLLATE NOCASE
                """, (new_thai_word.strip(), now, en))

        conn.commit()
        affected = cur.rowcount
        conn.close()
        logger.info(f"update_word '{identifier}': สำเร็จ {affected} แถว")
        return affected > 0
    except Exception as e:
        logger.error(f"update_word error: {e}")
        return False


def add_word(english_word: str, thai_word: str, note: str = "") -> int:
    """
    เพิ่มคำทับศัพท์คำใหม่เข้าฐานข้อมูล
    คืนค่า ID ที่ถูกสร้าง
    """
    try:
        conn = get_connection()
        cur = conn.cursor()

        # ตรวจสอบว่ามีคำภาษาอังกฤษนี้อยู่แล้วหรือไม่
        cur.execute(
            "SELECT id FROM transliterations WHERE english_word = ? COLLATE NOCASE",
            (english_word.strip(),)
        )
        existing = cur.fetchone()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if existing:
            rec_id = existing[0]
            cur.execute("""
                UPDATE transliterations 
                SET thai_word = ?, note = ?, is_active = 1, updated_at = ?
                WHERE id = ?
            """, (thai_word.strip(), note.strip(), now, rec_id))
            conn.commit()
            conn.close()
            return rec_id

        # หา MAX ID เพื่อรันต่อ
        cur.execute("SELECT MAX(id) FROM transliterations")
        max_id_row = cur.fetchone()
        next_id = (max_id_row[0] or 0) + 1

        cur.execute("""
            INSERT INTO transliterations (id, english_word, thai_word, note, category, is_active, created_at, updated_at)
            VALUES (?, ?, ?, ?, 'transliteration', 1, ?, ?)
        """, (next_id, english_word.strip(), thai_word.strip(), note.strip(), now, now))

        conn.commit()
        conn.close()
        logger.info(f"add_word: เพิ่ม '{english_word}' (ID: {next_id}) สำเร็จ")
        return next_id
    except Exception as e:
        logger.error(f"add_word error: {e}")
        return -1


def delete_word(identifier) -> bool:
    """ลบหรือ Deactivate คำศัพท์"""
    try:
        conn = get_connection()
        cur = conn.cursor()
        if isinstance(identifier, int) or (isinstance(identifier, str) and identifier.isdigit()):
            cur.execute("DELETE FROM transliterations WHERE id = ?", (int(identifier),))
        else:
            cur.execute("DELETE FROM transliterations WHERE english_word = ? COLLATE NOCASE", (str(identifier).strip(),))
        conn.commit()
        affected = cur.rowcount
        conn.close()
        return affected > 0
    except Exception as e:
        logger.error(f"delete_word error: {e}")
        return False


# ============================================================
# 3. ฟังก์ชันอัปเดตยกชุด / รีเซ็ต (Bulk Replace & Reset)
# ============================================================

def bulk_replace_entries(entries: List[Dict]) -> int:
    """
    แทนที่ข้อมูลทั้งหมดด้วยชุดข้อมูลใหม่ (Bulk Replace)
    entries: list of dict มีคีย์ english_word, thai_word, (note)
    """
    if not entries:
        return 0

    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM transliterations")

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        to_insert = []
        for i, e in enumerate(entries, 1):
            to_insert.append((
                e.get("id", i),
                e.get("english_word", "").strip(),
                e.get("thai_word", "").strip(),
                e.get("note", "").strip(),
                e.get("category", "transliteration"),
                1,
                now,
                now,
            ))

        cur.executemany("""
            INSERT INTO transliterations (id, english_word, thai_word, note, category, is_active, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, to_insert)

        conn.commit()
        conn.close()
        logger.info(f"bulk_replace_entries สำเร็จ {len(to_insert)} รายการ")
        return len(to_insert)
    except Exception as e:
        logger.error(f"bulk_replace_entries error: {e}")
        return 0


def import_from_uploaded_file(file_bytes: bytes, filename: str) -> Tuple[bool, str, int]:
    """
    นำเข้าและแทนที่ฐานข้อมูลจากไฟล์ที่ user อัปโหลด (CSV, XLSX, หรือ JSON)
    Returns: (success: bool, message: str, count: int)
    """
    try:
        entries = []
        lower_name = filename.lower()

        if lower_name.endswith(".csv"):
            # ลอง decode utf-8-sig ก่อน
            text = None
            for enc in ["utf-8-sig", "utf-8", "cp874", "tis-620"]:
                try:
                    text = file_bytes.decode(enc)
                    break
                except Exception:
                    continue
            if not text:
                return False, "ไม่สามารถอ่านการเข้ารหัสของไฟล์ CSV ได้ (รองรับ UTF-8 / TIS-620)", 0

            reader = csv.reader(io.StringIO(text))
            rows = list(reader)
            if len(rows) < 2:
                return False, "ไฟล์ไม่มีข้อมูลแถว", 0

            # หาตำแหน่งคอลัมน์จาก header
            header = [h.strip().lower() for h in rows[0]]
            en_idx, th_idx, note_idx, id_idx = 1, 2, 3, 0

            for i, h in enumerate(header):
                if any(k in h for k in ["english", "ภาษาอังกฤษ", "en", "incorrect_word"]):
                    en_idx = i
                elif any(k in h for k in ["thai", "ภาษาไทย", "ทับศัพท์", "th", "correct_word"]):
                    th_idx = i
                elif any(k in h for k in ["note", "หมายเหตุ"]):
                    note_idx = i
                elif any(k in h for k in ["ลำดับ", "id", "no"]):
                    id_idx = i

            for idx, r in enumerate(rows[1:], 1):
                if len(r) > max(en_idx, th_idx):
                    en = r[en_idx].strip()
                    th = r[th_idx].strip()
                    note = r[note_idx].strip() if len(r) > note_idx else ""
                    row_id = int(r[id_idx]) if (len(r) > id_idx and r[id_idx].isdigit()) else idx
                    if en or th:
                        entries.append({
                            "id": row_id,
                            "english_word": en,
                            "thai_word": th,
                            "note": note,
                        })

        elif lower_name.endswith(".xlsx"):
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                return False, "ไฟล์ Excel ไม่มีข้อมูล", 0

            header = [str(h).strip().lower() if h else "" for h in rows[0]]
            en_idx, th_idx, note_idx, id_idx = 1, 2, 3, 0
            for i, h in enumerate(header):
                if any(k in h for k in ["english", "ภาษาอังกฤษ", "en", "incorrect_word"]):
                    en_idx = i
                elif any(k in h for k in ["thai", "ภาษาไทย", "ทับศัพท์", "th", "correct_word"]):
                    th_idx = i
                elif any(k in h for k in ["note", "หมายเหตุ"]):
                    note_idx = i
                elif any(k in h for k in ["ลำดับ", "id", "no"]):
                    id_idx = i

            for idx, r in enumerate(rows[1:], 1):
                if len(r) > max(en_idx, th_idx):
                    en = str(r[en_idx]).strip() if r[en_idx] is not None else ""
                    th = str(r[th_idx]).strip() if r[th_idx] is not None else ""
                    note = str(r[note_idx]).strip() if len(r) > note_idx and r[note_idx] is not None else ""
                    row_id = int(r[id_idx]) if len(r) > id_idx and str(r[id_idx]).isdigit() else idx
                    if en or th:
                        entries.append({
                            "id": row_id,
                            "english_word": en,
                            "thai_word": th,
                            "note": note,
                        })

        elif lower_name.endswith(".json"):
            data = json.loads(file_bytes.decode("utf-8"))
            if isinstance(data, list):
                for i, item in enumerate(data, 1):
                    en = item.get("english_word") or item.get("english") or item.get("en") or ""
                    th = item.get("thai_word") or item.get("thai") or item.get("th") or ""
                    note = item.get("note", "")
                    if en or th:
                        entries.append({
                            "id": item.get("id", i),
                            "english_word": en,
                            "thai_word": th,
                            "note": note,
                        })
        else:
            return False, f"ชนิดไฟล์ไม่รองรับ: {filename} (รองรับ .csv, .xlsx, .json)", 0

        if not entries:
            return False, "ไม่พบข้อมูลคำศัพท์ที่สามารถนำเข้าได้", 0

        count = bulk_replace_entries(entries)
        return True, f"นำเข้าข้อมูลและอัปเดตฐานข้อมูลสำเร็จ ({count} คำ)", count
    except Exception as e:
        logger.error(f"import_from_uploaded_file error: {e}")
        return False, f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}", 0


def reset_to_default_database() -> int:
    """รีเซ็ตฐานข้อมูลกลับเป็นค่าเริ่มต้น 1,561 คำเดิมจากตัวสร้าง"""
    try:
        from database_builder import main as build_main
        build_main()
        return get_total_count()
    except Exception as e:
        logger.error(f"reset_to_default_database error: {e}")
        return 0


# ============================================================
# 4. Engine Transliteration & Misspelling Matcher
# ============================================================

# รูปแบบคำผิดยอดนิยมในภาษาไทยทางการที่มักสะกดคลาดเคลื่อน
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
    "อัปเดท": "อัปเดต",
    "อัพเกรด": "อัปเกรด",
    "อัปเกรท": "อัปเกรด",
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


def get_engine_lookup_db() -> Dict[str, Dict]:
    """
    สร้างและคืนค่า lookup dictionary ที่ครอบคลุมสำหรับ Checker Engine:
    1. คำภาษาอังกฤษ 1,561 คำ -> แนะนำคำทับศัพท์ไทย
    2. คำทับศัพท์ไทยที่สะกดผิดยอดนิยม -> แนะนำคำทับศัพท์ทางการ
    """
    lookup = {}

    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT id, english_word, thai_word, note FROM transliterations WHERE is_active = 1")
        for row in cur.fetchall():
            en = row["english_word"].strip()
            th = row["thai_word"].strip()
            note = row["note"].strip()

            if en and th:
                # 1. แนะนำเมื่อพบคำภาษาอังกฤษ (ให้เขียนเป็นทับศัพท์ไทย)
                lookup[en] = {
                    "correct": th,
                    "note": f"คำภาษาอังกฤษ ควรใช้คำทับศัพท์ทางการ: \"{th}\"{' (' + note + ')' if note else ''}",
                    "type": "transliteration",
                    "is_english": True,
                }
        conn.close()
    except Exception as e:
        logger.error(f"get_engine_lookup_db error: {e}")

    # 2. ผนวกคำทับศัพท์ภาษาไทยที่สะกดผิด
    for wrong_th, correct_th in COMMON_MISPELLING_MAP.items():
        if wrong_th not in lookup:
            lookup[wrong_th] = {
                "correct": correct_th,
                "note": f"คำทับศัพท์ที่มักเขียนผิดตามหลักราชบัณฑิตยสภา/วุฒิสภา (ที่ถูกต้องคือ \"{correct_th}\")",
                "type": "transliteration",
                "is_english": False,
            }

    return lookup
