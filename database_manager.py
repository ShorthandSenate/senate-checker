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
    # ============================================================
    # 1. คำทับศัพท์ตามคู่มือคำทับศัพท์ทางการวุฒิสภา (๑,๕๖๑ คำ / กิจวุฒิ)
    #    และคำทับศัพท์เทคโนโลยี/ดิจิทัลที่มักสะกดผิด
    # ============================================================
    # account -> แอ็กเคานต์
    "แอคเคาท์": "แอ็กเคานต์",
    "แอคเค้าท์": "แอ็กเคานต์",
    "แอคเค้า": "แอ็กเคานต์",
    "แอ็คเคาท์": "แอ็กเคานต์",
    "แอกเคาท์": "แอ็กเคานต์",

    # smart -> สมาร์ต
    "สมาร์ท": "สมาร์ต",
    "สมาร์ทโฟน": "สมาร์ตโฟน",
    "สมาร์ทซิตี": "สมาร์ตซิตี",
    "สมาร์ทฟาร์มเมอร์": "สมาร์ตฟาร์มเมอร์",

    # app / application -> แอป / แอปพลิเคชัน
    "แอพ": "แอป",
    "แอพพลิเคชัน": "แอปพลิเคชัน",
    "แอพพลิเคชั่น": "แอปพลิเคชัน",
    "แอปพลิเคชั่น": "แอปพลิเคชัน",

    # digital -> ดิจิทัล
    "ดิจิตอล": "ดิจิทัล",
    "ดิจิตอลวอลเล็ต": "ดิจิทัลวอลเล็ต",

    # block / blockchain -> บล็อก / บล็อกเชน
    "บล็อค": "บล็อก",
    "บล็อคเชน": "บล็อกเชน",
    "บล็อกเชนจ์": "บล็อกเชน",

    # update / upgrade -> อัปเดต / อัปเกรด
    "อัพเดท": "อัปเดต",
    "อัพเดต": "อัปเดต",
    "อัปเดท": "อัปเดต",
    "อัพเกรด": "อัปเกรด",
    "อัปเกรท": "อัปเกรด",

    # link -> ลิงก์
    "ลิงค์": "ลิงก์",
    "ลิ้งค์": "ลิงก์",
    "ลิ้ง": "ลิงก์",

    # click -> คลิก
    "คลิ๊ก": "คลิก",
    "คลิกส์": "คลิก",

    # email -> อีเมล
    "อีเมล์": "อีเมล",
    "อีเมลล์": "อีเมล",

    # platform -> แพลตฟอร์ม
    "แพลทฟอร์ม": "แพลตฟอร์ม",
    "แพลตฟอร์ม์": "แพลตฟอร์ม",

    # website -> เว็บไซต์
    "เว็ปไซต์": "เว็บไซต์",
    "เวบไซต์": "เว็บไซต์",
    "เวปไซต์": "เว็บไซต์",
    "เว็บไซด์": "เว็บไซต์",
    "เว็ปไซด์": "เว็บไซต์",

    # charge -> ชาร์จ (ใช้รูปผสม เพื่อป้องกันการตรวจชนคำว่า พายชาร์ต / โฟลว์ชาร์ต / ฟลิปชาร์ต)
    "ชาร์ตไฟ": "ชาร์จไฟ",
    "ชาร์ตแบต": "ชาร์จแบต",
    "ชาร์ตแบตเตอรี่": "ชาร์จแบตเตอรี่",
    "ชาร์ตแบตเตอรี": "ชาร์จแบตเตอรี่",
    "สายชาร์ต": "สายชาร์จ",
    "ที่ชาร์ต": "ที่ชาร์จ",
    "หัวชาร์ต": "หัวชาร์จ",
    "แท่นชาร์ต": "แท่นชาร์จ",
    "ชาร์ตพลังงาน": "ชาร์จพลังงาน",

    # post -> โพสต์
    "โพส": "โพสต์",
    "โพสท์": "โพสต์",
    "โพสต์ท": "โพสต์",

    # computer -> คอมพิวเตอร์
    "คอมพิวเต้อร์": "คอมพิวเตอร์",
    "คอมพิวเตอร์์": "คอมพิวเตอร์",

    # YouTube -> ยูทูบ
    "ยูทูป": "ยูทูบ",
    "ยูทู๊บ": "ยูทูบ",

    # project -> โพรเจกต์
    "โปรเจค": "โพรเจกต์",
    "โปรเจ็ค": "โพรเจกต์",
    "โพรเจค": "โพรเจกต์",
    "โปรเจกต์": "โพรเจกต์",
    "โปรเจ็กต์": "โพรเจ็กต์",

    # function -> ฟังก์ชัน
    "ฟังก์ชั่น": "ฟังก์ชัน",
    "ฟังชั่น": "ฟังก์ชัน",
    "ฟังค์ชั่น": "ฟังก์ชัน",

    # graphic -> กราฟิก
    "กราฟฟิก": "กราฟิก",
    "กราฟฟิค": "กราฟิก",
    "กราฟิค": "กราฟิก",

    # sensor / center -> เซนเซอร์ / เซนเตอร์
    "เซ็นเซอร์": "เซนเซอร์",
    "เซ็นเตอร์": "เซนเตอร์",

    # check -> เช็ก
    "เช็คอิน": "เช็กอิน",
    "เช็ค": "เช็ก",

    # software / soft power -> ซอฟต์แวร์ / ซอฟต์พาวเวอร์
    "ซอฟท์แวร์": "ซอฟต์แวร์",
    "ซอฟท์พาวเวอร์": "ซอฟต์พาวเวอร์",
    "ซอฟพาวเวอร์": "ซอฟต์พาวเวอร์",

    # start -> สตาร์ต
    "สตาร์ท": "สตาร์ต",
    "สตาร์ทอัพ": "สตาร์ตอัป",

    # sticker / step / status -> สติกเกอร์ / สเตป / สเตตัส
    "สติ๊กเกอร์": "สติกเกอร์",
    "สเต็ป": "สเตป",
    "สเตตัสส์": "สเตตัส",

    # quota -> โควตา
    "โควต้า": "โควตา",

    # taxi -> แท็กซี
    "แท็กซี่": "แท็กซี",
    "แท๊กซี่": "แท็กซี",

    # point -> พอยต์
    "พอยท์": "พอยต์",
    "พ้อยท์": "พอยต์",

    # party -> ปาร์ตี
    "ปาร์ตี้": "ปาร์ตี",
    "ปาร์ตี้ลิสต์": "ปาร์ตีลิสต์",

    # Shopee -> ช้อปปี / shopping -> ช้อปปิง
    "ช็อปปี": "ช้อปปี",
    "ช็อปปิ้ง": "ช้อปปิง",
    "ช้อปปิ้ง": "ช้อปปิง",

    # internet -> อินเทอร์เน็ต
    "อินเตอร์เน็ต": "อินเทอร์เน็ต",
    "อินเตอร์เนต": "อินเทอร์เน็ต",
    "อินเทอร์เนต": "อินเทอร์เน็ต",

    # comment -> คอมเมนต์
    "คอมเม้นต์": "คอมเมนต์",
    "คอมเม้นท์": "คอมเมนต์",
    "คอมเมนท์": "คอมเมนต์",

    # like -> ไลก์
    "ไลค์": "ไลก์",
    "ไล้ค์": "ไลก์",

    # Facebook -> เฟซบุ๊ก
    "เฟสบุ๊ค": "เฟซบุ๊ก",
    "เฟซบุ๊ค": "เฟซบุ๊ก",
    "เฟสบุ๊ก": "เฟซบุ๊ก",

    # TikTok -> ติ๊กต็อก
    "ติ๊กต๊อก": "ติ๊กต็อก",
    "ติ๊กตอก": "ติ๊กต็อก",

    # Google -> กูเกิล
    "กูเกิ้ล": "กูเกิล",

    # Microsoft -> ไมโครซอฟต์
    "ไมโครซอฟท์": "ไมโครซอฟต์",

    # streaming -> สตรีมมิง
    "สตรีมมิ่ง": "สตรีมมิง",

    # content -> คอนเทนต์
    "คอนเท้นต์": "คอนเทนต์",
    "คอนเทนท์": "คอนเทนต์",

    # workshop -> เวิร์กชอป
    "เวิร์คช็อป": "เวิร์กชอป",
    "เวิร์คชอป": "เวิร์กชอป",
    "เวิร์กช็อป": "เวิร์กชอป",

    # chat -> แชต
    "แชท": "แชต",
    "แช็ต": "แชต",

    # login -> ล็อกอิน
    "ล็อคอิน": "ล็อกอิน",
    "ล๊อคอิน": "ล็อกอิน",

    # server -> เซิร์ฟเวอร์
    "เซิฟเวอร์": "เซิร์ฟเวอร์",
    "เซิร์ฟเว่อร์": "เซิร์ฟเวอร์",

    # download / upload -> ดาวน์โหลด / อัปโหลด
    "ดาวโหลด": "ดาวน์โหลด",
    "ดาวน์โหลต": "ดาวน์โหลด",
    "อัพโหลด": "อัปโหลด",
    "อัปโหลต": "อัปโหลด",

    # profile / style -> โปรไฟล์ / สไตล์
    "โปรฟาย": "โปรไฟล์",
    "โปรฟายล์": "โปรไฟล์",
    "สไตร์": "สไตล์",

    # guarantee -> การันตี
    "การันตี๋": "การันตี",

    # bakery / calorie / ice cream / series -> เบเกอรี / แคลอรี / ไอศกรีม / ซีรีส์
    "เบเกอรี่": "เบเกอรี",
    "แคลอรี่": "แคลอรี",
    "ไอศครีม": "ไอศกรีม",
    "ไอติม": "ไอศกรีม",
    "ซีรีย์": "ซีรีส์",
    "ซีรี่ส์": "ซีรีส์",

    # ============================================================
    # 2. ศัพท์กฎหมาย นิติบัญญัติ และงานสารบรรณวุฒิสภา
    #    (ตรวจสอบอย่างเคร่งครัดตามข้อบังคับและพจนานุกรมราชบัณฑิตยสภา)
    # ============================================================
    # กลุ่ม: กฎหมาย (ใช้ ฎ ชฎา)
    "กฏหมาย": "กฎหมาย",
    "กฏกระทรวง": "กฎกระทรวง",
    "กฏข้อบังคับ": "กฎข้อบังคับ",
    "กฏเกณฑ์": "กฎเกณฑ์",
    "กฏระเบียบ": "กฎระเบียบ",

    # กลุ่ม: ปรากฏ (ใช้ ฏ ปฏัก)
    "ปรากฎ": "ปรากฏ",
    "ปรากฎการณ์": "ปรากฏการณ์",

    # กลุ่ม: มงกุฎ (ใช้ ฎ ชฎา)
    "มงกุฏ": "มงกุฎ",

    # กลุ่ม: ปฏิบัติ / ปฏิรูป (ใช้ ฏ ปฏัก)
    "ปฎิบัติ": "ปฏิบัติ",
    "ปฎิบัติการ": "ปฏิบัติการ",
    "ปฎิบัติงาน": "ปฏิบัติงาน",
    "ปฎิบัติหน้าที่": "ปฏิบัติหน้าที่",
    "ปฎิรูป": "ปฏิรูป",

    # กลุ่ม: แปรญัตติ (กระบวนการนิติบัญญัติในสภา ใช้ 'แปรญัตติ' ห้ามใช้ 'แปล')
    "แปลญัตติ": "แปรญัตติ",
    "การแปลญัตติ": "การแปรญัตติ",
    "คำขอแปลญัตติ": "คำขอแปรญัตติ",

    # กลุ่ม: สังเกต (ไม่มีสระอุ)
    "สังเกตุ": "สังเกต",
    "ข้อสังเกตุ": "ข้อสังเกต",
    "สังเกตุการณ์": "สังเกตการณ์",

    # กลุ่ม: โอกาส / อากาศ
    "โอกาศ": "โอกาส",
    "อากาส": "อากาศ",

    # กลุ่ม: รสชาติ / ศีรษะ
    "รสชาด": "รสชาติ",
    "ศรีษะ": "ศีรษะ",

    # กลุ่ม: ลายเซ็น / เซ็นชื่อ (ไม่ใช้คำเดี่ยว 'เซ็นต์' เพราะจะตรวจชนคำว่า 'เปอร์เซ็นต์')
    "ลายเซ็นต์": "ลายเซ็น",
    "เซ็นต์ชื่อ": "เซ็นชื่อ",
    "เซ็นต์รับ": "เซ็นรับ",
    "เซ็นต์เอกสาร": "เซ็นเอกสาร",
    "เซ็นต์ยินยอม": "เซ็นยินยอม",
    "เซ็นต์สัญญา": "เซ็นสัญญา",
    "เซ็นต์อนุมัติ": "เซ็นอนุมัติ",

    # เปอร์เซ็นต์ (ตามคู่มือคำทับศัพท์ทางการวุฒิสภา ให้เขียนติดกัน ไม่เว้นวรรค)
    "เปอร์ เซ็นต์": "เปอร์เซ็นต์",
    "เปอร์เซนต์": "เปอร์เซ็นต์",

    # กลุ่ม: ผาสุก / บริสุทธิ์
    "ผาสุข": "ผาสุก",
    "บริสุทธิ": "บริสุทธิ์",

    # กลุ่ม: รัฐสภาและกระบวนการประชุม
    "ยุทธสาสตร์": "ยุทธศาสตร์",
    "อภิบาย": "อภิปราย",
    "พิจารนา": "พิจารณา",
    "งบประมาน": "งบประมาณ",
    "นวัฒกรรม": "นวัตกรรม",
    "วัฒนธรรม์": "วัฒนธรรม",
    "มติที่ประขุม": "มติที่ประชุม",
    "ข้อบังคัง": "ข้อบังคับ",
    "รัฐธรรมนูน": "รัฐธรรมนูญ",
    "ข้าราชการณ": "ข้าราชการ",
    "ประทานวุฒิสภา": "ประธานวุฒิสภา",
    "รองประทาน": "รองประธาน",
    "ราชกิจจานุเบกษาา": "ราชกิจจานุเบกษา",
    "สัมฤทธิ์ผล": "สัมฤทธิผล",
    "เอกฉันทร์": "เอกฉันท์",

    # กลุ่ม: องค์ประชุม vs ห้องประชุม (แก้ปัญหาการตัดคำหลุด 'ห้ องประชุม' และเจาะจงเฉพาะองค์ประชุม)
    "ห้ องประชุม": "ห้องประชุม",
    "ในห้ องประชุม": "ในห้องประชุม",
    "ครบองประชุม": "ครบองค์ประชุม",
    "ไม่ครบองประชุม": "ไม่ครบองค์ประชุม",
    "นับองประชุม": "นับองค์ประชุม",
    "ตรวจองประชุม": "ตรวจองค์ประชุม",
    "เป็นองประชุม": "เป็นองค์ประชุม",
    "กรรมาธิการณ": "กรรมาธิการ",
    "กรรมาธิกาณ": "กรรมาธิการ",
    "กฤษฏีกา": "กฤษฎีกา",
    "กฤษฎีกาา": "กฤษฎีกา",
    "พระราชกฤษฎีกาา": "พระราชกฤษฎีกา",
    "กระทรวงสาธรณสุข": "กระทรวงสาธารณสุข",

    # กลุ่ม: สัมมนา / ประสิทธิผล
    "อนุญาติ": "อนุญาต",
    "สัมนา": "สัมมนา",
    "สัมนาการ": "สัมมนาการ",
    "สัมานา": "สัมมนา",
    "สัมนากร": "สัมมนากร",
    "ประสิทธิ์ผล": "ประสิทธิผล",
    "ประสิทธิผลล": "ประสิทธิผล",
    "ประสิทธิ์ภาพ": "ประสิทธิภาพ",
    "อนุกรรมาธิกาณ": "อนุกรรมาธิการ",
    "อนุกรมาธิการ": "อนุกรรมาธิการ",
    "ผ้อำนวยการ": "ผู้อำนวยการ",
    "เลขาธิการณ": "เลขาธิการ",
    "ผู้บริหาล": "ผู้บริหาร",
    "บริหาล": "บริหาร",
    "ผลิตภัณท์": "ผลิตภัณฑ์",
    "ผลิตภัน": "ผลิตภัณฑ์",
    "ผลิดภัณฑ์": "ผลิตภัณฑ์",
    "ส่ิงแวดล้อม": "สิ่งแวดล้อม",
    "ส่ิงที่": "สิ่งที่",
    "ปัจจุบับ": "ปัจจุบัน",
    "แลละ": "และ",
    "ผูกพันธ์": "ผูกพัน",
    "อานิสงฆ์": "อานิสงส์",
    "โลกาภิวัฒน์": "โลกาภิวัตน์",
    "กิติมศักดิ์": "กิตติมศักดิ์",
    "วิพากวิจารณ์": "วิพากษ์วิจารณ์",
    "เกมส์": "เกม",
    "ผัดวันประกันพรุ่ง": "ผลัดวันประกันพรุ่ง",
    "ลำใย": "ลำไย",
    "กระเพรา": "กะเพรา",
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
