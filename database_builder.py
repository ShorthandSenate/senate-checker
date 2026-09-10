# ============================================================
# database_builder.py - แปลงและรวบรวมฐานข้อมูลวุฒิสภาฉบับสมบูรณ์
# (ไม่รวมคำย่อ ฝฝ ตามที่ผู้ใช้แจ้ง)
# 1. คำทับศัพท์ทางการ 1,561 รายการ
# 2. ทำเนียบสมาชิกวุฒิสภา (สว. 2567) 200 ท่าน (วรรคใหญ่ ๒ เคาะ)
# 3. คำศัพท์ทางการ/กระทรวง/คณะกรรมาธิการ/หน่วยงานรัฐ 348 รายการ
# 4. ข้อความเหตุการณ์การประชุม 24 รายการ
# ============================================================

import os
import sys
import re
import html
import sqlite3
import json
import csv
from datetime import datetime

# บังคับใช้ UTF-8 สำหรับ stdout บน Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
UPLOAD_DIR = r"C:\Users\NB_DELL\.gemini\antigravity\brain\4b8de4dd-ec8a-4d81-9eec-2d0116824a0c\.user_uploaded"

HTML_TRANSLIT_PATH = os.path.join(UPLOAD_DIR, "media_1789014062063.html")
HTML_SENATORS_PATH = os.path.join(UPLOAD_DIR, "media_1789023123751.html")
HTML_TERMS_PATH = os.path.join(UPLOAD_DIR, "media_1789023123757.html")

SQLITE_DB_PATH = os.path.join(DATA_DIR, "vocab.db")


# ============================================================
# 1. Parsing Functions
# ============================================================

def parse_transliterations(html_path: str) -> list:
    """อ่านและสกัดแถวข้อมูลคำทับศัพท์ 1,561 รายการ"""
    if not os.path.exists(html_path):
        raise FileNotFoundError(f"ไม่พบไฟล์: {html_path}")

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    tr_blocks = re.findall(r"<tr[^>]*>(.*?)</tr>", content, re.DOTALL)
    entries = []

    for tr in tr_blocks:
        tds = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.DOTALL)
        clean_tds = [html.unescape(re.sub(r"<[^>]+>", "", td)).strip() for td in tds]

        if len(clean_tds) >= 4 and clean_tds[1].isdigit():
            num = int(clean_tds[1])
            en = clean_tds[2].strip()
            th = clean_tds[3].strip()
            note = clean_tds[4].strip() if len(clean_tds) > 4 else ""

            if "พจนานุกรม" in note:
                note = "พจนานุกรม"
            elif "ศัพท์บัญญัติ" in note or "ศััพท์บัญญัติ" in note:
                note = "ศัพท์บัญญัติ"
            elif "หลักนิยม" in note or "หลัักนิยม" in note or "หลักนิิยม" in note:
                note = "คำตามหลักนิยม"

            if en or th:
                entries.append({
                    "id": num,
                    "english_word": en,
                    "thai_word": th,
                    "note": note,
                    "category": "transliteration",
                })

    return entries


def parse_senators_and_events(html_path: str) -> tuple:
    """
    อ่านและสกัดสมาชิกวุฒิสภา 200 คน และข้อความเหตุการณ์การประชุม 24 รายการ
    (ไม่รวมคำย่อ ฝฝ)
    """
    if not os.path.exists(html_path):
        raise FileNotFoundError(f"ไม่พบไฟล์: {html_path}")

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    tr_blocks = re.findall(r"<tr[^>]*>(.*?)</tr>", content, re.DOTALL)
    senators = []
    events = []

    known_titles = [
        "ผู้ช่วยศาสตราจารย์พิเศษ", "ผู้ช่วยศาสตราจารย์", "รองศาสตราจารย์",
        "ศาสตราจารย์พิเศษ", "ศาสตราจารย์เกียรติคุณ", "ศาสตราจารย์",
        "พลตำรวจเอก", "พลตำรวจโท", "พลตำรวจตรี", "พันตำรวจเอก", "พันตำรวจโท", "พันตำรวจตรี", "ร้อยตำรวจเอก",
        "พลอากาศเอก", "พลอากาศโท", "พลอากาศตรี",
        "พลเรือเอก", "พลเรือโท", "พลเรือตรี", "นาวาตรี",
        "พลเอก", "พลโท", "พลตรี", "พันเอกหญิง", "พันเอก", "พันโท", "พันตรี", "ร้อยเอก",
        "ว่าที่ร้อยตรี", "ว่าที่พันตรี", "นายแพทย์",
        "คุณหญิง", "หม่อมหลวง", "นาย", "นางสาว", "นาง"
    ]

    military_police_ranks = {
        "พลตำรวจเอก", "พลตำรวจโท", "พลตำรวจตรี", "พันตำรวจเอก", "พันตำรวจโท", "พันตำรวจตรี",
        "ร้อยตำรวจเอก", "ร้อยตำรวจโท", "ร้อยตำรวจตรี",
        "พลอากาศเอก", "พลอากาศโท", "พลอากาศตรี", "นาวาอากาศเอก", "นาวาอากาศโท", "นาวาอากาศตรี",
        "พลเรือเอก", "พลเรือโท", "พลเรือตรี", "นาวาเอก", "นาวาโท", "นาวาตรี",
        "พลเอก", "พลโท", "พลตรี", "พันเอกหญิง", "พันเอก", "พันโท", "พันตรี",
        "ร้อยเอก", "ร้อยโท", "ร้อยตรี", "ว่าที่ร้อยตรี", "ว่าที่พันตรี",
    }

    for tr in tr_blocks:
        tds = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.DOTALL)
        clean = [html.unescape(re.sub(r"<[^>]+>", "", t)).strip() for t in tds]

        if len(clean) >= 5:
            col_img = clean[1]
            col_no = clean[2]
            # clean[3] คือคำย่อ ฝฝ -> ไม่นำมาใช้
            col_full = clean[4]

            # ข้อมูล สว. 200 ท่าน
            if col_no.startswith("เลขที่ "):
                seat_str = col_no.replace("เลขที่ ", "").strip()
                seat_num = int(seat_str) if seat_str.isdigit() else len(senators) + 1

                # แยกชื่อตัว ชื่อกลาง นามสกุล
                parts = [p.strip() for p in re.split(r"\s{2,}", col_full) if p.strip()]

                title = ""
                first_name = ""
                middle_name = ""
                last_name = ""

                first_part = parts[0] if len(parts) > 0 else col_full
                if len(parts) == 2:
                    last_name = parts[1]
                elif len(parts) >= 3:
                    middle_name = parts[1]
                    last_name = parts[2]

                # ดึง title จาก first_part
                matched_title = ""
                rem_name = first_part
                for t in sorted(known_titles, key=len, reverse=True):
                    if first_part.startswith(t):
                        matched_title = t
                        rem_name = first_part[len(t):].strip()
                        break

                title = matched_title
                first_name = rem_name

                # จัดรูปแบบมาตรฐานทางการ:
                # 1. คำนำหน้านามที่เป็นยศ: ต้องพิมพ์ห่างกับชื่อตัว เช่น พลเอก  เกรียงไกร (วรรค ๒ เคาะ)
                # 2. คำนำหน้านามบุคคลธรรมดา/วิชาการ: พิมพ์ติดกับชื่อตัว เช่น นายกมล
                # 3. ระหว่างชื่อตัว (ชื่อกลาง) และนามสกุล: เว้นวรรคใหญ่ ๒ เคาะ เสมอ
                rank_space = "  " if title in military_police_ranks else ""
                if middle_name:
                    full_name_official = f"{title}{rank_space}{first_name}  {middle_name}  {last_name}".strip()
                else:
                    full_name_official = f"{title}{rank_space}{first_name}  {last_name}".strip()

                senators.append({
                    "id": seat_num,
                    "seat_no": col_no,
                    "title": title,
                    "first_name": first_name,
                    "middle_name": middle_name,
                    "last_name": last_name,
                    "full_name_official": full_name_official,
                    "image_url": col_img,
                })

            # ข้อมูลข้อความเหตุการณ์การประชุม 24 รายการ (ไม่รวมคำย่อ ฝฝ)
            elif col_full and not col_no.startswith("เลขที่") and not col_full.startswith("คำเต็ม"):
                # ตัด row header ที่ไม่ใช่ข้อความ
                if "(" in col_full or "ผู้ปฏิบัติหน้าที่" in col_full:
                    events.append({
                        "id": len(events) + 1,
                        "description": col_full,
                    })

    return senators, events


def parse_parliament_entities(html_path: str) -> list:
    """
    อ่านและสกัดคลังชื่อหน่วยงาน กระทรวง คณะกรรมาธิการ และคำทางการรัฐสภา
    (ไม่รวมคำย่อ ฝฝ)
    """
    if not os.path.exists(html_path):
        raise FileNotFoundError(f"ไม่พบไฟล์: {html_path}")

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    tr_blocks = re.findall(r"<tr[^>]*>(.*?)</tr>", content, re.DOTALL)
    entities = []
    seen = set()

    for tr in tr_blocks:
        tds = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.DOTALL)
        clean = [html.unescape(re.sub(r"<[^>]+>", "", t)).strip() for t in tds]

        if len(clean) >= 4:
            # clean[2] คือคำย่อ ฝฝ -> ไม่เอา
            full = clean[3].strip()
            if full and full not in ["คำเต็มอัตโนมัติ", "C", "B", "A"] and len(full) > 2 and full not in seen:
                seen.add(full)

                # จัดหมวดหมู่
                category = "หน่วยงานรัฐสภา"
                if "กระทรวง" in full:
                    category = "กระทรวง"
                elif "คณะกรรมาธิการ" in full:
                    category = "คณะกรรมาธิการ"
                elif "คณะกรรมการ" in full:
                    category = "คณะกรรมการ"
                elif "สำนักงาน" in full or "สำนัก" in full:
                    category = "สำนักงาน"

                entities.append({
                    "id": len(entities) + 1,
                    "entity_name": full,
                    "category": category,
                })

    return entities


# ============================================================
# 2. SQLite Database Builder
# ============================================================

def build_all_sqlite_tables(
    transliterations: list,
    senators: list,
    events: list,
    entities: list,
    db_path: str
):
    """สร้างตารางทั้งหมดในฐานข้อมูล SQLite พร้อม Index โดยไม่มีคำย่อ ฝฝ"""
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # 1. ตารางคำทับศัพท์
    cur.execute("DROP TABLE IF EXISTS transliterations")
    cur.execute("""
        CREATE TABLE transliterations (
            id INTEGER PRIMARY KEY,
            english_word TEXT NOT NULL,
            thai_word TEXT NOT NULL,
            note TEXT DEFAULT '',
            category TEXT DEFAULT 'transliteration',
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.executemany("""
        INSERT INTO transliterations (id, english_word, thai_word, note, category)
        VALUES (:id, :english_word, :thai_word, :note, :category)
    """, transliterations)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_translit_en ON transliterations(english_word COLLATE NOCASE)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_translit_th ON transliterations(thai_word)")

    # 2. ตารางสมาชิกวุฒิสภา 200 คน (วรรคใหญ่ ๒ เคาะ)
    cur.execute("DROP TABLE IF EXISTS senators")
    cur.execute("""
        CREATE TABLE senators (
            id INTEGER PRIMARY KEY,
            seat_no TEXT,
            title TEXT,
            first_name TEXT,
            middle_name TEXT,
            last_name TEXT,
            full_name_official TEXT NOT NULL,
            image_url TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.executemany("""
        INSERT INTO senators (id, seat_no, title, first_name, middle_name, last_name, full_name_official, image_url)
        VALUES (:id, :seat_no, :title, :first_name, :middle_name, :last_name, :full_name_official, :image_url)
    """, senators)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_senator_full ON senators(full_name_official)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_senator_first ON senators(first_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_senator_last ON senators(last_name)")

    # 3. ตารางข้อความเหตุการณ์การประชุม 24 รายการ
    cur.execute("DROP TABLE IF EXISTS meeting_events")
    cur.execute("""
        CREATE TABLE meeting_events (
            id INTEGER PRIMARY KEY,
            description TEXT NOT NULL,
            is_active INTEGER DEFAULT 1
        )
    """)
    cur.executemany("""
        INSERT INTO meeting_events (id, description)
        VALUES (:id, :description)
    """, events)

    # 4. ตารางชื่อหน่วยงาน/คำศัพท์ทางการรัฐสภา 348 รายการ
    cur.execute("DROP TABLE IF EXISTS parliament_entities")
    cur.execute("DROP TABLE IF EXISTS parliament_terms")
    cur.execute("""
        CREATE TABLE parliament_entities (
            id INTEGER PRIMARY KEY,
            entity_name TEXT NOT NULL,
            category TEXT DEFAULT 'parliament_entity',
            is_active INTEGER DEFAULT 1
        )
    """)
    cur.executemany("""
        INSERT INTO parliament_entities (id, entity_name, category)
        VALUES (:id, :entity_name, :category)
    """, entities)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_entity_name ON parliament_entities(entity_name)")

    conn.commit()
    conn.close()
    print(f"✓ บันทึกข้อมูลเข้า SQLite สำเร็จ: {db_path}")
    print(f"   • คำทับศัพท์ทางการ: {len(transliterations):,} คำ")
    print(f"   • ทำเนียบสมาชิกวุฒิสภา: {len(senators):,} ท่าน (วรรคใหญ่ ๒ เคาะ)")
    print(f"   • ข้อความเหตุการณ์การประชุม: {len(events):,} รายการ")
    print(f"   • คำศัพท์/หน่วยงานทางการรัฐสภา: {len(entities):,} รายการ (ไม่มีคำย่อ ฝฝ)")


# ============================================================
# 3. Export CSV & Excel Files
# ============================================================

def export_all_files(transliterations: list, senators: list, events: list, entities: list):
    """ส่งออกข้อมูลทั้งหมดเป็น CSV (UTF-8 BOM) และ Excel สวยงาม"""
    # 1. Transliterations
    csv_trans = os.path.join(DATA_DIR, "transliterations.csv")
    with open(csv_trans, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ลำดับที่", "ภาษาอังกฤษ", "คำทับศัพท์ภาษาไทย", "หมายเหตุ"])
        for e in transliterations:
            w.writerow([e["id"], e["english_word"], e["thai_word"], e["note"]])

    csv_engine = os.path.join(DATA_DIR, "vocab_db.csv")
    with open(csv_engine, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["incorrect_word", "correct_word", "note", "type"])
        for e in transliterations:
            w.writerow([e["english_word"], e["thai_word"], f"คำทับศัพท์ทางการวุฒิสภา{' (' + e['note'] + ')' if e['note'] else ''}", "transliteration"])

    wb_trans = Workbook()
    ws_trans = wb_trans.active
    ws_trans.title = "คำทับศัพท์วุฒิสภา"
    ws_trans.append(["ลำดับที่", "ภาษาอังกฤษ", "คำทับศัพท์ภาษาไทย", "หมายเหตุ"])
    for e in transliterations:
        ws_trans.append([e["id"], e["english_word"], e["thai_word"], e["note"]])
    wb_trans.save(os.path.join(DATA_DIR, "transliterations.xlsx"))

    # 2. Senators 200 ท่าน
    csv_sen = os.path.join(DATA_DIR, "senators.csv")
    with open(csv_sen, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ลำดับที่", "เลขที่", "คำนำหน้า", "ชื่อตัว", "ชื่อกลาง", "นามสกุล", "ชื่อเต็มทางการ (วรรค ๒ เคาะ)", "รูปภาพ"])
        for s in senators:
            w.writerow([s["id"], s["seat_no"], s["title"], s["first_name"], s["middle_name"], s["last_name"], s["full_name_official"], s["image_url"]])

    wb_sen = Workbook()
    ws_sen = wb_sen.active
    ws_sen.title = "สมาชิกวุฒิสภา 2567"
    ws_sen.append(["ลำดับที่", "เลขที่", "คำนำหน้า", "ชื่อตัว", "ชื่อกลาง", "นามสกุล", "ชื่อเต็มทางการ (วรรค ๒ เคาะ)"])
    for s in senators:
        ws_sen.append([s["id"], s["seat_no"], s["title"], s["first_name"], s["middle_name"], s["last_name"], s["full_name_official"]])
    wb_sen.save(os.path.join(DATA_DIR, "senators.xlsx"))

    # 3. Parliament Entities & Events
    with open(os.path.join(DATA_DIR, "parliament_entities.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ลำดับ", "ชื่อหน่วยงาน/องค์กร/คำทางการ", "หมวดหมู่"])
        for ent in entities:
            w.writerow([ent["id"], ent["entity_name"], ent["category"]])

    with open(os.path.join(DATA_DIR, "meeting_events.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ลำดับ", "ข้อความเหตุการณ์การประชุม"])
        for ev in events:
            w.writerow([ev["id"], ev["description"]])

    print("✓ ส่งออกไฟล์ CSV และ Excel สำเร็จครบทุกตาราง")


def main():
    print("🚀 กำลังรวบรวมและสร้างฐานข้อมูลวุฒิสภา (ตัดคำย่อ ฝฝ ออกทั้งหมด)...")
    os.makedirs(DATA_DIR, exist_ok=True)

    transliterations = parse_transliterations(HTML_TRANSLIT_PATH)
    senators, events = parse_senators_and_events(HTML_SENATORS_PATH)
    entities = parse_parliament_entities(HTML_TERMS_PATH)

    build_all_sqlite_tables(transliterations, senators, events, entities, SQLITE_DB_PATH)
    export_all_files(transliterations, senators, events, entities)
    print("🎉 สร้างและผนวกฐานข้อมูลสำเร็จครบ 100%!")


if __name__ == "__main__":
    main()
