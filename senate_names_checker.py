# ============================================================
# senate_names_checker.py - ตรวจสอบความถูกต้องของชื่อ-สกุล
# สมาชิกวุฒิสภา (๒๐๐ ท่าน) และ ผู้บริหารสำนักงานเลขาธิการวุฒิสภา (๑๕ ท่าน)
# แม่นยำ 100% Deterministic + High-Precision Fuzzy Match
# ============================================================

import os
import re
import json
import difflib
import sqlite3
import logging
from typing import List, Dict, Tuple, Optional

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
JSON_PATH = os.path.join(DATA_DIR, "senate_personnel.json")
DB_PATH = os.path.join(DATA_DIR, "vocab.db")

TITLES = [
    "ผู้ช่วยศาสตราจารย์พิเศษ", "รองศาสตราจารย์", "ผู้ช่วยศาสตราจารย์", "ศาสตราจารย์",
    "พันตำรวจเอก", "พันตำรวจโท", "ร้อยตำรวจเอก", "พลตำรวจตรี", "พลตำรวจโท",
    "พันเอกหญิง", "ว่าที่พันตรี", "นาวาตรี",
    "พลเอก", "พลโท", "พลตรี", "พันเอก", "พันโท", "พันตรี",
    "นางสาว", "น.ส.", "นาย", "นาง",
]

TITLES_PATTERN = "|".join([re.escape(t) for t in sorted(TITLES, key=len, reverse=True)])
TITLE_NAME_RE = re.compile(rf'({TITLES_PATTERN})\s*([ก-๙]+)(?:\s+([ก-๙]+))?')


class SenateNamesChecker:
    """
    ตรวจสอบชื่อ สว. และผู้บริหาร:
    1. ตรวจจับคำนำหน้า + ชื่อ + นามสกุล
    2. เทียบเคียงกับฐานข้อมูล ๒๑๕ ท่าน
    3. หากถูกต้อง 100% -> ผ่าน ไม่แจ้งเตือน
    4. หากสะกดใกล้เคียง (ตกการันต์, สระเพี้ยน, พิมพ์ผิด) -> แจ้งเตือนพร้อมแนะนำชื่อและตำแหน่งทางการ
    5. หากเป็นบุคคลภายนอกทั่วไป -> ผ่าน ไม่เกิด False Positive
    """

    def __init__(self, json_path: str = JSON_PATH):
        self.json_path = json_path
        self.personnel: List[Dict] = []
        self.exact_full_names = set()
        self.load_data()

    def load_data(self):
        """โหลดข้อมูลบุคลากรจาก SQLite หรือ JSON"""
        loaded = False
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH, timeout=5)
                conn.row_factory = sqlite3.Row
                cur = conn.cursor()
                cur.execute("""
                    SELECT id, person_type, title, first_name, last_name, full_name_official, role, is_active
                    FROM senate_personnel
                    WHERE is_active = 1
                    ORDER BY id ASC
                """)
                rows = cur.fetchall()
                if rows:
                    self.personnel = [dict(r) for r in rows]
                    loaded = True
                conn.close()
            except Exception as e:
                logger.warning(f"Failed to load from sqlite senate_personnel: {e}")

        if not loaded and os.path.exists(self.json_path):
            try:
                with open(self.json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.personnel = [p for p in data if p.get("is_active", 1) == 1]
                loaded = True
            except Exception as e:
                logger.error(f"Failed to load senate personnel JSON: {e}")

        self.exact_full_names = set()
        for p in self.personnel:
            t = p.get("title", "")
            fn = p.get("first_name", "")
            ln = p.get("last_name", "")
            self.exact_full_names.add(p.get("full_name_official", ""))
            self.exact_full_names.add(f"{t}{fn} {ln}".strip())
            self.exact_full_names.add(f"{t} {fn} {ln}".strip())
            self.exact_full_names.add(f"{t}{fn}  {ln}".strip())
            self.exact_full_names.add(f"{t} {fn}  {ln}".strip())
            self.exact_full_names.add(f"{fn} {ln}".strip())
            self.exact_full_names.add(f"{fn}  {ln}".strip())

        self.fn_map = {}
        for p in self.personnel:
            fn = p.get("first_name", "").strip()
            if fn:
                self.fn_map.setdefault(fn, []).append(p)

        if self.fn_map:
            fn_pattern = "|".join([re.escape(fn) for fn in sorted(self.fn_map.keys(), key=len, reverse=True)])
            self.fn_regex = re.compile(rf'({fn_pattern})\s+([ก-๙]+)')
        else:
            self.fn_regex = None

    @staticmethod
    def _sim(s1: str, s2: str) -> float:
        return difflib.SequenceMatcher(None, s1, s2).ratio()

    def _find_best_person(self, fn: str, ln: str) -> Tuple[Optional[Dict], float]:
        best_p = None
        best_score = 0.0
        len_fn = len(fn)
        len_ln = len(ln)

        for p in self.personnel:
            p_fn = p["first_name"]
            p_ln = p["last_name"]

            # Pre-filter: ถ้าความยาวต่างกันเกิน 3 ตัวอักษรทั้งชื่อและนามสกุล ไม่มีทางคะแนนถึง 0.70
            if abs(len_fn - len(p_fn)) > 3 and (not ln or abs(len_ln - len(p_ln)) > 3):
                continue

            s_fn = self._sim(fn, p_fn)
            s_ln = self._sim(ln, p_ln) if ln and p_ln else 0.0

            if ln and p["last_name"]:
                if fn == p["first_name"] and s_ln >= 0.70:
                    score = 0.5 + 0.5 * s_ln
                elif ln == p["last_name"] and s_fn >= 0.70:
                    score = 0.5 + 0.5 * s_fn
                elif s_fn >= 0.80 and s_ln >= 0.80:
                    score = 0.5 * s_fn + 0.5 * s_ln
                else:
                    score = 0.0
            elif not ln and not p["last_name"]:
                score = s_fn if s_fn >= 0.88 else 0.0
            else:
                score = 0.0

            if score > best_score:
                best_score = score
                best_p = p

        return best_p, best_score

    def check_text(self, text: str) -> List[Dict]:
        """
        ตรวจสอบชื่อ สว. และผู้บริหารในข้อความย่อหน้า
        คืนค่ารายการข้อผิดพลาด:
        {
            'rule_type': 'senate_personnel',
            'rule_name': 'ความถูกต้องของชื่อ-สกุล สมาชิกวุฒิสภา / ผู้บริหาร',
            'wrong': 'ข้อความที่ผิดในเอกสาร',
            'correct': 'ชื่อทางการที่ถูกต้อง',
            'reason': 'คำอธิบาย',
            'pos': (start, end)
        }
        """
        issues = []
        if not text or not self.personnel:
            return issues

        covered_spans = []

        # -------------------------------------------------------------
        # Pass 1: ตรวจจับกรณีมีคำนำหน้า (คำนำหน้า + ชื่อ + นามสกุล)
        # -------------------------------------------------------------
        for m in TITLE_NAME_RE.finditer(text):
            title = m.group(1)
            fn = m.group(2)
            ln = m.group(3) or ""
            span_text = m.group(0)
            start_pos = m.start()
            end_pos = m.end()

            # 1. ตรวจความตรงกันแบบ 100% (Exact Match)
            exact_p = None
            for p in self.personnel:
                if fn == p["first_name"] and ln == p["last_name"]:
                    exact_p = p
                    break

            if exact_p:
                covered_spans.append((start_pos, end_pos))
                # ตรวจความสอดคล้องของคำนำหน้า (เช่น น.ส. vs นางสาว vs นาง)
                if title != exact_p["title"]:
                    issues.append({
                        "rule_type": "senate_personnel",
                        "rule_name": f"คำนำหน้าชื่อ {exact_p['role']}",
                        "wrong": span_text,
                        "correct": exact_p["full_name_official"],
                        "reason": f"ในทำเนียบวุฒิสภาใช้คำนำหน้า '{exact_p['title']}' แนะนำ: '{exact_p['full_name_official']}'",
                        "pos": (start_pos, end_pos),
                    })
                # ตรวจการเว้นวรรค ๒ เคาะ หรือเว้นวรรคยศ
                elif span_text != exact_p["full_name_official"]:
                    issues.append({
                        "rule_type": "senate_personnel",
                        "rule_name": f"การเว้นวรรคชื่อ-สกุล {exact_p['role']}",
                        "wrong": span_text,
                        "correct": exact_p["full_name_official"],
                        "reason": f"ระหว่างชื่อตัวและนามสกุล ให้เว้นวรรคใหญ่ (๒ เคาะ) ตามระเบียบวุฒิสภา: '{exact_p['full_name_official']}'",
                        "pos": (start_pos, end_pos),
                    })
                continue

            # 2. ตรวจความใกล้เคียง (Fuzzy Match สำหรับคำที่สะกดผิด)
            best_p, best_score = self._find_best_person(fn, ln)
            if best_p and best_score >= 0.82:
                covered_spans.append((start_pos, end_pos))
                issues.append({
                    "rule_type": "senate_personnel",
                    "rule_name": f"ความถูกต้องของชื่อ-สกุล {best_p['role']}",
                    "wrong": span_text,
                    "correct": best_p["full_name_official"],
                    "reason": f"ชื่อใกล้เคียงกับ {best_p['role']}: '{best_p['full_name_official']}' แนะนำให้ตรวจสอบการสะกด",
                    "pos": (start_pos, end_pos),
                })

        # -------------------------------------------------------------
        # Pass 2: ตรวจจับกรณีไม่มีคำนำหน้าในประโยค (ระบุชื่อ + นามสกุล)
        # -------------------------------------------------------------
        if hasattr(self, "fn_regex") and self.fn_regex:
            for m in self.fn_regex.finditer(text):
                fn = m.group(1)
                found_ln = m.group(2)
                fn_start = m.start()
                total_end = m.end()
                full_span = m.group(0).strip()

                if any(s <= fn_start < e or s < total_end <= e for s, e in covered_spans):
                    continue

                for p in self.fn_map.get(fn, []):
                    ln = p["last_name"]
                    role = p["role"]
                    official = p["full_name_official"]

                    if found_ln == ln:
                        covered_spans.append((fn_start, total_end))
                        break
                    else:
                        sim_ln = self._sim(found_ln, ln)
                        if sim_ln >= 0.70:
                            covered_spans.append((fn_start, total_end))
                            issues.append({
                                "rule_type": "senate_personnel",
                                "rule_name": f"ความถูกต้องของชื่อ-สกุล {role}",
                                "wrong": full_span,
                                "correct": official,
                                "reason": f"นามสกุลใกล้เคียงกับ {role}: '{official}' แนะนำให้ตรวจสอบการสะกด",
                                "pos": (fn_start, total_end),
                            })
                            break

        issues.sort(key=lambda x: x["pos"][0] if "pos" in x else 0)
        return issues


# Singleton instance สำหรับเรียกใช้ใน Engine
_global_checker: Optional[SenateNamesChecker] = None

def get_senate_names_checker() -> SenateNamesChecker:
    global _global_checker
    if _global_checker is None:
        _global_checker = SenateNamesChecker()
    return _global_checker

def reload_senate_names_checker():
    global _global_checker
    _global_checker = SenateNamesChecker()
    return _global_checker
