# ============================================================
# database_manager.py - ระบบจัดการฐานข้อมูลและคลังข้อมูลวุฒิสภาฉบับสมบูรณ์
# 1. ฐานข้อมูลคำทับศัพท์ทางการ (1,561 คำ)
# 2. ทำเนียบสมาชิกวุฒิสภา (สว. 2567) 200 ท่าน
# 3. ข้อความเหตุการณ์การประชุม 24 รายการ
# 4. คำย่อและคำศัพท์ทางการรัฐสภา 352 รายการ
# 5. กฎระเบียบสำนักกรรมาธิการ ๓
# ============================================================

import os
import re
import csv
import sys
import json
import sqlite3
import logging
from typing import Dict, List, Optional, Tuple

# บังคับใช้ UTF-8 สำหรับ stdout
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
SQLITE_DB_PATH = os.path.join(DATA_DIR, "vocab.db")

CSV_TRANSLIT_PATH = os.path.join(DATA_DIR, "transliterations.csv")
CSV_ENGINE_PATH = os.path.join(DATA_DIR, "vocab_db.csv")
CSV_SENATORS_PATH = os.path.join(DATA_DIR, "senators.csv")
CSV_TERMS_PATH = os.path.join(DATA_DIR, "parliament_terms.csv")
CSV_EVENTS_PATH = os.path.join(DATA_DIR, "meeting_events.csv")


def get_db_connection() -> Optional[sqlite3.Connection]:
    """เปิด connection ไปยัง SQLite DB ถ้ามี"""
    if os.path.exists(SQLITE_DB_PATH):
        try:
            conn = sqlite3.connect(SQLITE_DB_PATH, timeout=10)
            conn.row_factory = sqlite3.Row
            return conn
        except Exception as e:
            logger.error(f"sqlite3 connect error: {e}")
    return None


def init_db_if_needed():
    """ตรวจสอบว่าไฟล์ฐานข้อมูลมีครบหรือไม่ ถ้าไม่มีให้ข้ามไป (ใช้ CSV fallback แทน)"""
    try:
        if not os.path.exists(SQLITE_DB_PATH) or os.path.getsize(SQLITE_DB_PATH) == 0:
            # บน Cloud อาจไม่มี HTML source files สำหรับ builder
            # ให้ใช้ CSV fallback แทนโดยไม่ต้อง build
            logger.warning("vocab.db not found — will use CSV fallback")
    except Exception as e:
        logger.error(f"init_db_if_needed error: {e}")


# เรียก init อัตโนมัติเมื่อ import โมดูล (safe — จะไม่ crash)
try:
    init_db_if_needed()
except Exception:
    pass


# ============================================================
# 1. จัดการคำทับศัพท์ (Transliterations 1,561 คำ)
# ============================================================

# ============================================================
# ตารางคำทับศัพท์ภาษาไทยที่มักสะกดผิดตามหลักราชบัณฑิตยสภา / วุฒิสภา
# รูปแบบ: "คำผิด": "คำถูก"
# สามารถเพิ่มคำใหม่ได้ที่นี่ได้เลย — ระบบจะอ่านและตรวจจากตารางนี้อัตโนมัติ
# ============================================================
COMMON_MISPELLING_MAP = {
    # --- คำทับศัพท์อิเล็กทรอนิกส์และเทคโนโลยี ---
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
    "สติ๊กเกอร์": "สติกเกอร์",
    "สเต็ป": "สเตป",
    "โควต้า": "โควตา",
    "แท็กซี่": "แท็กซี",
    "แท๊กซี่": "แท็กซี",
    "พอยท์": "พอยต์",
    "พ้อยท์": "พอยต์",
    "ปาร์ตี้": "ปาร์ตี",
    "ปาร์ตี้ลิสต์": "ปาร์ตีลิสต์",
    # เพิ่ม ช้อปปี (Shopee) ตาม Request
    "ช็อปปี": "ช้อปปี",

    # ============================================================
    # คำผิดภาษาไทยทั่วไปที่มักพบในรายงานการประชุม
    # (พิมพ์ตก / เติมไม่ครบ / สะกดผิด / สับสนพยัญชนะ)
    # ============================================================

    # กลุ่ม: อนุญาต (มักใส่ ต เป็น ติ)
    "อนุญาติ": "อนุญาต",

    # กลุ่ม: สัมมนา
    "สัมนา": "สัมมนา",
    "สัมนาการ": "สัมมนาการ",
    "สัมานา": "สัมมนา",
    "สัมนากร": "สัมมนากร",

    # กลุ่ม: ประสิทธิ-
    "ประสิทธิ์ผล": "ประสิทธิผล",
    "ประสิทธิผลล": "ประสิทธิผล",
    "ประสิทธิ์ภาพ": "ประสิทธิภาพ",

    # กลุ่ม: กรรมาธิการ (มักสะกดผิด)
    "อนุกรรมาธิกาณ": "อนุกรรมาธิการ",
    "อนุกรมาธิการ": "อนุกรรมาธิการ",
    "กรรมาธิกาณ": "กรรมาธิการ",

    # กลุ่ม: คำทั่วไปที่มักสะกดผิด
    "ผ้อำนวยการ": "ผู้อำนวยการ",
    "เลขาธิการณ": "เลขาธิการ",
    "ผู้บริหาล": "ผู้บริหาร",
    "บริหาล": "บริหาร",

    # กลุ่ม: สับสน ณ / น / ร / ล
    "ผลิตภัณท์": "ผลิตภัณฑ์",
    "ผลิตภัน": "ผลิตภัณฑ์",
    "ผลิดภัณฑ์": "ผลิตภัณฑ์",

    # กลุ่ม: พิมพ์ผิดสระ
    "ส่ิงแวดล้อม": "สิ่งแวดล้อม",
    "ส่ิงที่": "สิ่งที่",
    "ปัจจุบับ": "ปัจจุบัน",

    # กลุ่ม: คำทางกฎหมาย/รัฐสภา
    "พระราชกฤษฎีกาา": "พระราชกฤษฎีกา",

    # กลุ่ม: คำสันธาน
    "แลละ": "และ",

    # กลุ่ม: หน่วยงานรัฐ
    "กระทรวงสาธรณสุข": "กระทรวงสาธารณสุข",

    # กลุ่ม: คำไทยและคำทับศัพท์ที่มักเขียนผิดตามพจนานุกรมราชบัณฑิตยสภา
    "ผูกพันธ์": "ผูกพัน",
    "เซ็นต์ชื่อ": "เซ็นชื่อ",
    "เซ็นต์": "เซ็น",
    "อานิสงฆ์": "อานิสงส์",
    "โลกาภิวัฒน์": "โลกาภิวัตน์",
    "กิติมศักดิ์": "กิตติมศักดิ์",
    "วิพากวิจารณ์": "วิพากษ์วิจารณ์",
    "เกมส์": "เกม",
}


def get_total_count() -> int:
    """นับจำนวนคำทับศัพท์ทางการทั้งหมดในระบบ"""
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM transliterations WHERE is_active = 1")
            cnt = cur.fetchone()[0]
            conn.close()
            return cnt
        except Exception:
            conn.close()

    # Fallback to CSV
    if os.path.exists(CSV_TRANSLIT_PATH):
        try:
            with open(CSV_TRANSLIT_PATH, "r", encoding="utf-8-sig") as f:
                return max(0, sum(1 for _ in f) - 1)
        except Exception:
            pass
    return 1561


def get_engine_lookup_db() -> Dict[str, Dict]:
    """
    สร้าง lookup dictionary สำหรับตรวจคำทับศัพท์และคำภาษาอังกฤษ:
    1. คำภาษาอังกฤษ 1,561 คำ -> แนะนำคำทับศัพท์ไทยทางการ
    2. คำทับศัพท์ไทยที่สะกดผิดยอดนิยม -> แนะนำคำทับศัพท์ที่ถูกต้อง
    """
    lookup = {}

    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT english_word, thai_word, note FROM transliterations WHERE is_active = 1")
            for row in cur.fetchall():
                en = row["english_word"].strip()
                th = row["thai_word"].strip()
                note = row["note"].strip()

                if en and th:
                    lookup[en] = {
                        "correct": th,
                        "note": f"คำภาษาอังกฤษ ควรใช้คำทับศัพท์ทางการ: \"{th}\"{(' (' + note + ')') if note else ''}",
                        "type": "transliteration",
                        "is_english": True,
                    }
            conn.close()
        except Exception as e:
            logger.error(f"get_engine_lookup_db from sqlite error: {e}")
            if conn:
                conn.close()

    # Fallback to CSV if lookup is empty
    if not lookup and os.path.exists(CSV_TRANSLIT_PATH):
        try:
            with open(CSV_TRANSLIT_PATH, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    en = (row.get("ภาษาอังกฤษ") or row.get("english_word") or "").strip()
                    th = (row.get("คำทับศัพท์ภาษาไทย") or row.get("thai_word") or "").strip()
                    note = row.get("หมายเหตุ") or row.get("note") or ""
                    if en and th:
                        lookup[en] = {
                            "correct": th,
                            "note": f"คำภาษาอังกฤษ ควรใช้คำทับศัพท์ทางการ: \"{th}\"",
                            "type": "transliteration",
                            "is_english": True,
                        }
        except Exception as e:
            logger.error(f"get_engine_lookup_db from csv error: {e}")

    return lookup


def add_word(english_word: str, thai_word: str, note: str = "") -> bool:
    """เพิ่มหรืออัปเดตคำทับศัพท์เฉพาะคำ"""
    conn = get_db_connection()
    if not conn:
        return False
    try:
        cur = conn.cursor()
        now = "CURRENT_TIMESTAMP"
        cur.execute("SELECT id FROM transliterations WHERE english_word = ? COLLATE NOCASE", (english_word.strip(),))
        row = cur.fetchone()
        if row:
            cur.execute("""
                UPDATE transliterations 
                SET thai_word = ?, note = ?, is_active = 1, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (thai_word.strip(), note.strip(), row["id"]))
        else:
            cur.execute("SELECT MAX(id) FROM transliterations")
            max_id = (cur.fetchone()[0] or 0) + 1
            cur.execute("""
                INSERT INTO transliterations (id, english_word, thai_word, note, category, is_active)
                VALUES (?, ?, ?, ?, 'transliteration', 1)
            """, (max_id, english_word.strip(), thai_word.strip(), note.strip()))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        logger.error(f"add_word error: {e}")
        if conn:
            conn.close()
        return False


def reset_to_default_database() -> int:
    """รีเซ็ตฐานข้อมูลกลับเป็นค่าเริ่มต้น"""
    try:
        from database_builder import main as build_main
        build_main()
        return get_total_count()
    except Exception as e:
        logger.error(f"reset_to_default_database error: {e}")
        return 0


# ============================================================
# 2. จัดการข้อมูลสมาชิกวุฒิสภา 200 ท่าน (Senators Roster)
# ============================================================

def get_senators_count() -> int:
    """นับจำนวน สว. ในฐานข้อมูล"""
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM senators WHERE is_active = 1")
            cnt = cur.fetchone()[0]
            conn.close()
            return cnt
        except Exception:
            conn.close()
    return 200


def get_senators_lookup() -> List[Dict]:
    """
    ดึงรายชื่อ สว. ทั้งหมด 200 ท่าน เพื่อให้ Engine นำไปตรวจสอบ:
    - ชื่อนามสกุลทางการที่ถูกต้อง (เว้นวรรคใหญ่ ๒ เคาะ)
    - รูปแบบที่มักสะกดผิดหรือเว้นวรรคเพียง ๑ เคาะ เพื่อใช้ตรวจจับ
    """
    senators = []
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, seat_no, title, first_name, middle_name, last_name, full_name_official, image_url 
                FROM senators 
                WHERE is_active = 1
                ORDER BY id ASC
            """)
            for r in cur.fetchall():
                senators.append(dict(r))
            conn.close()
            return senators
        except Exception as e:
            logger.error(f"get_senators_lookup from sqlite error: {e}")
            if conn:
                conn.close()

    # Fallback to CSV
    if os.path.exists(CSV_SENATORS_PATH):
        try:
            with open(CSV_SENATORS_PATH, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    senators.append({
                        "id": int(r.get("ลำดับที่", 0)),
                        "seat_no": r.get("เลขที่", ""),
                        "title": r.get("คำนำหน้า", ""),
                        "first_name": r.get("ชื่อตัว", ""),
                        "middle_name": r.get("ชื่อกลาง", ""),
                        "last_name": r.get("นามสกุล", ""),
                        "full_name_official": r.get("ชื่อเต็มทางการ (วรรค ๒ เคาะ)", ""),
                    })
        except Exception as e:
            logger.error(f"get_senators_lookup from csv error: {e}")

    return senators


# ============================================================
# 3. คำศัพท์ทางการ/กระทรวง/คณะกรรมาธิการ/หน่วยงานรัฐ (349 รายการ)
# ============================================================

def get_parliament_entities_lookup() -> List[str]:
    """ดึงรายชื่อกระทรวง คณะกรรมาธิการ และหน่วยงานรัฐสภา (ไม่มีคำย่อ ฝฝ)"""
    entities = []
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT entity_name FROM parliament_entities WHERE is_active = 1")
            for r in cur.fetchall():
                entities.append(r["entity_name"].strip())
            conn.close()
            return entities
        except Exception as e:
            logger.error(f"get_parliament_entities_lookup sqlite error: {e}")
            if conn:
                conn.close()

    # Fallback to CSV
    csv_entities_path = os.path.join(DATA_DIR, "parliament_entities.csv")
    if os.path.exists(csv_entities_path):
        try:
            with open(csv_entities_path, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    name = r.get("ชื่อหน่วยงาน/องค์กร/คำทางการ", "").strip()
                    if name:
                        entities.append(name)
        except Exception:
            pass

    return entities


# ============================================================
# 4. ข้อความเหตุการณ์การประชุม 24 รายการ (Meeting Events)
# ============================================================

def get_meeting_events_lookup() -> List[str]:
    """ดึงข้อความเหตุการณ์การประชุม 24 รายการ (ไม่มีคำย่อ ฝฝ)"""
    events = []
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT description FROM meeting_events WHERE is_active = 1")
            for r in cur.fetchall():
                events.append(r["description"].strip())
            conn.close()
            return events
        except Exception as e:
            logger.error(f"get_meeting_events_lookup sqlite error: {e}")
            if conn:
                conn.close()

    # Fallback to CSV
    if os.path.exists(CSV_EVENTS_PATH):
        try:
            with open(CSV_EVENTS_PATH, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    desc = r.get("ข้อความเหตุการณ์การประชุม", "").strip()
                    if desc:
                        events.append(desc)
        except Exception:
            pass

    return events


# ============================================================
# 5. กฎระเบียบเฉพาะสำนักกรรมาธิการ ๓ (Editorial Rules Definition)
# ============================================================

SENATE_EDITORIAL_RULES = [
    {
        "id": "recreation_word",
        "type": "vocabulary",
        "wrong_pattern": r"สันทนาการ",
        "correct_replacement": "นันทนาการ",
        "reason": "คำว่า \"สันทนาการ\" ถูกยกเลิกโดยราชบัณฑิตยสภาแล้ว ให้เปลี่ยนเป็น \"นันทนาการ\" ตามระเบียบสำนักกรรมาธิการ ๓",
    },
    {
        "id": "senator_from_province",
        "type": "spacing",
        "wrong_pattern": r"(สมาชิกวุฒิสภา)\s+(จาก(?:จังหวัด|อำเภอ))",
        "correct_replacement": r"\1\2",
        "reason": "คำว่า \"จาก\" ให้พิมพ์ติดกับคำว่า \"สมาชิกวุฒิสภา\" โดยไม่ต้องเว้นวรรค ตามระเบียบสำนักกรรมาธิการ ๓",
    },
    {
        "id": "thung_to_jueng",
        "type": "spelling",
        "wrong_pattern": r"(เหตุใด|ทำไม|ด้วยเหตุดังกล่าว)ถึง(เกิด|ทำให้|มี|เป็น)",
        "correct_replacement": r"\1จึง\2",
        "reason": "คำว่า \"ถึง\" ในความหมายแสดงผลลัพธ์ ให้เปลี่ยนเป็น \"จึง\" ตามระเบียบสำนักกรรมาธิการ ๓",
    },
    {
        "id": "phraram_bridge_river",
        "type": "vocabulary",
        "wrong_pattern": r"(สะพานพระราม|เขื่อนพระราม)ที่\s*([๔๕๖๗๘๙\d]+)",
        "correct_replacement": r"\1 \2",
        "reason": "ชื่อสะพานและเขื่อนพระราม ทุกแห่งไม่มีคำว่า \"ที่\" (เช่น สะพานพระราม ๕, สะพานพระราม ๘) ตามระเบียบสำนักกรรมาธิการ ๓",
    },
    {
        "id": "phraram9_road",
        "type": "vocabulary",
        "wrong_pattern": r"ถนนพระรามที่\s*([๙9])",
        "correct_replacement": r"ถนนพระราม \1",
        "reason": "ถนนพระราม ๙ ไม่มีคำว่า \"ที่\" ตามระเบียบสำนักกรรมาธิการ ๓ (ยกเว้น พระรามที่ ๑ ถึง ๖ มีคำว่า ที่)",
    },
    {
        "id": "compound_nospace",
        "type": "spacing",
        "wrong_pattern": r"(บำเหน็จ|ธำรง|บำรุง|ฉกชิง|อุปโภค)\s+(บำนาญ|รักษา|วิ่งราว|บริโภค)",
        "correct_replacement": r"\1\2",
        "reason": "คำที่ใช้ควบคู่กัน (บำเหน็จบำนาญ, ธำรงรักษา, บำรุงรักษา, ฉกชิงวิ่งราว, อุปโภคบริโภค) ให้พิมพ์ติดกันไม่ต้องเว้นวรรค",
    },
    {
        "id": "no_space_before_kue",
        "type": "spacing",
        "wrong_pattern": r"([^\s\d\(\[\{]+)\s+(คือ)(?!\s)",
        "correct_replacement": r"\1\2",
        "reason": "คำว่า \"คือ\" ในสำนวนทั่วไปไม่ต้องเว้นวรรคหน้า ให้พิมพ์ติดกับคำหน้า ตามระเบียบสำนักกรรมาธิการ ๓",
    },
    {
        "id": "page_reference_word",
        "type": "vocabulary",
        "wrong_pattern": r"หน้าที่\s*([\d๑-๙]+)\b",
        "correct_replacement": r"หน้า \1",
        "reason": "การพูดอ้างถึงเนื้อหาในเอกสารรายงาน ที่มีคำว่า \"หน้าที่\" ให้แก้ไขเป็น \"หน้า\" ตามระเบียบสำนักกรรมาธิการ ๓",
    },
    {
        "id": "such_as_and_so_on",
        "type": "vocabulary",
        "wrong_pattern": r"(เช่น\s*[^,\.\n]+?)\s*เป็นต้น",
        "correct_replacement": r"\1",
        "reason": "คำว่า \"เช่น\" กับ \"เป็นต้น\" ในประโยคเดียวกันไม่ใช้คู่กัน ให้ตัด \"เป็นต้น\" ออก ตามระเบียบสำนักกรรมาธิการ ๓",
    },
    # --- กฎเพิ่มเติมจากคู่มือการจัดทำรายงานการประชุมวุฒิสภา (๖๔ หน้า) ---
    {
        "id": "video_term_spelling",
        "type": "spelling",
        "wrong_pattern": r"(วีดีทัศน์|วิดีทัศน์|วีดีโอ|วิดิโอ)",
        "correct_replacement": "วีดิทัศน์",
        "reason": "คำว่า \"วีดิทัศน์\" ให้สะกดด้วย สระอี-ด-สระอิ (ครอบคลุมถึงคลิปและวิดีโอ) ตามคู่มือการจัดทำรายงานการประชุมวุฒิสภา หน้า ๒๓",
    },
    {
        "id": "statute_spacing_between",
        "type": "spacing",
        "wrong_pattern": r"(มาตรา\s*[\d๑-๙]+)[ \t](มาตรา\s*[\d๑-๙]+)",
        "correct_replacement": r"\1  \2",
        "reason": "ระหว่างมาตรา ให้เว้นวรรคใหญ่ (๒ เคาะ) เช่น มาตรา ๑  มาตรา ๒ ตามคู่มือการจัดทำรายงานการประชุมวุฒิสภา หน้า ๒๒",
    },
    {
        "id": "repeated_sound_nana",
        "type": "vocabulary",
        "wrong_pattern": r"นานาๆ",
        "correct_replacement": "ต่าง ๆ นานา",
        "reason": "คำ ๒ พยางค์ที่มีเสียงซ้ำกันไม่ควรใช้ไม้ยมก ให้เขียนเป็น \"ต่าง ๆ นานา\" ตามคู่มือการจัดทำรายงานการประชุมวุฒิสภา หน้า ๗",
    },
    {
        "id": "repeated_sound_chacha",
        "type": "vocabulary",
        "wrong_pattern": r"จะๆ",
        "correct_replacement": "จะจะ",
        "reason": "คำ ๒ พยางค์ที่มีเสียงซ้ำกันไม่ควรใช้ไม้ยมก ให้เขียนเป็น \"จะจะ\" ตามคู่มือการจัดทำรายงานการประชุมวุฒิสภา หน้า ๗",
    },
    {
        "id": "zero_votes_no_count",
        "type": "vocabulary",
        "wrong_pattern": r"(ไม่เห็นด้วย|งดออกเสียง)\s*([๐0])\s*คะแนน",
        "correct_replacement": r"\1 ไม่มี",
        "reason": "คะแนนที่เป็นศูนย์ ในผลการลงมติให้ใส่คำว่า \"ไม่มี\" และไม่ต้องใส่หน่วยนับคะแนน ตามคู่มือการจัดทำรายงานการประชุมวุฒิสภา หน้า ๒๖",
    },
    {
        "id": "law_term_prescribe_correct",
        "type": "vocabulary",
        "wrong_pattern": r"(รัฐธรรมนูญ|พระราชบัญญัติ|พระราชกำหนด)\s+(กำหนดว่า)",
        "correct_replacement": r"\1 บัญญัติว่า",
        "reason": "กฎหมายระดับรัฐธรรมนูญ พระราชบัญญัติ และพระราชกำหนด ให้ใช้คำว่า \"บัญญัติว่า\" (ไม่ใช่ กำหนดว่า) ตามคู่มือการจัดทำรายงานการประชุมวุฒิสภา หน้า ๕-๖",
    },
    {
        "id": "decree_term_prescribe_correct",
        "type": "vocabulary",
        "wrong_pattern": r"(พระราชกฤษฎีกา|กฎกระทรวง|ประกาศกระทรวง|ข้อบังคับ)\s+(บัญญัติว่า)",
        "correct_replacement": r"\1 กำหนดว่า",
        "reason": "กฎหมายลำดับรองระดับพระราชกฤษฎีกา กฎกระทรวง ประกาศกระทรวง และข้อบังคับ ให้ใช้คำว่า \"กำหนดว่า\" (ไม่ใช่ บัญญัติว่า) ตามคู่มือการจัดทำรายงานการประชุมวุฒิสภา หน้า ๕-๖",
    },
]


def get_senate_editorial_rules() -> List[Dict]:
    """ส่งออกรายการกฎเฉพาะของสำนักกรรมาธิการ ๓"""
    return SENATE_EDITORIAL_RULES
