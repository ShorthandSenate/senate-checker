# ============================================================
# rule_checker.py — High-Precision 100% Deterministic Rule-based Checker
# สำหรับรายงานการประชุมและเอกสารราชการวุฒิสภา
# ประกอบด้วย:
#   1. DictionaryChecker  — Trie Substring Engine (ไม่พึ่ง Tokenizer)
#   2. RegexChecker       — Pattern rules (ไม้ยมก, เอ.ไอ., ในขั้นกรรมาธิการ)
#   3. SenateDocumentChecker — รวมผลลัพธ์ List[dict]
# ============================================================
from __future__ import annotations

import re
import os
import logging
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

CONTEXT_WINDOW = 65


# ─────────────────────────────────────────────────────────────
# 1. _Trie: Pure-Python Trie สำหรับ Substring Matching
# ─────────────────────────────────────────────────────────────

class _Trie:
    """
    Trie-based substring matcher สำหรับภาษาไทย:
    - ค้นหาตรงจาก raw string ไม่ต้องตัดคำก่อน
    - Longest-match priority
    - Adjacent-keyword safe (ไม่ข้ามคำที่เขียนติดกัน)
    """
    _KW = "__kw__"

    def __init__(self):
        self.root: Dict = {}

    def add(self, word: str, payload: Any) -> None:
        node = self.root
        for ch in word:
            node = node.setdefault(ch, {})
        node[self._KW] = payload

    def extract(self, text: str) -> List[Tuple[Any, int, int]]:
        results: List[Tuple[Any, int, int]] = []
        n = len(text)
        idx = 0
        while idx < n:
            node = self.root
            match_payload = None
            match_end = idx
            cur = idx
            while cur < n and text[cur] in node:
                node = node[text[cur]]
                cur += 1
                if self._KW in node:
                    match_payload = node[self._KW]
                    match_end = cur
            if match_payload is not None:
                results.append((match_payload, idx, match_end))
                idx = match_end
            else:
                idx += 1
        return results


# ─────────────────────────────────────────────────────────────
# 2. DictionaryChecker
# ─────────────────────────────────────────────────────────────

class DictionaryChecker:
    """
    ตรวจหาคำผิดจาก Dictionary ด้วย Trie Substring Match
    ไม่พึ่งพา standard tokenizer — ตรวจจับได้ 100%
    """

    def __init__(self):
        self._trie = _Trie()
        self._registry: Dict[str, Tuple[str, str, str]] = {}

    def _register(
        self,
        wrong: str,
        correct: str,
        reason: str,
        rule_type: str = "misspelling",
        overwrite: bool = False,
    ) -> None:
        wrong = wrong.strip()
        correct = correct.strip()
        if not wrong or not correct or wrong == correct:
            return
        if wrong in self._registry and not overwrite:
            return
        self._registry[wrong] = (correct, reason, rule_type)
        self._trie.add(wrong, (wrong, correct, reason, rule_type))

    def load_dict(
        self,
        mapping: Dict[str, Any],
        rule_type: str = "misspelling",
        default_reason_prefix: str = "คำสะกดผิด",
        overwrite: bool = False,
    ) -> int:
        """
        โหลด mapping dict รูปแบบ:
          "คำผิด": "คำถูก"
          "คำผิด": ("คำถูก", "เหตุผล")
          "คำผิด": ("คำถูก", "เหตุผล", "rule_type")
        คืนจำนวนที่ ingest ได้จริง
        """
        count = 0
        for wrong, value in mapping.items():
            if isinstance(value, tuple):
                if len(value) >= 3:
                    correct, reason, rt = value[0], value[1], value[2]
                elif len(value) == 2:
                    correct, reason = value
                    rt = rule_type
                else:
                    correct = value[0]
                    reason = f"{default_reason_prefix}: ควรแก้เป็น '{correct}'"
                    rt = rule_type
            else:
                correct = str(value)
                reason = f"{default_reason_prefix}: ควรแก้เป็น '{correct}'"
                rt = rule_type
            self._register(wrong, correct, reason, rt, overwrite)
            count += 1
        return count

    def check_paragraph(
        self,
        para: Dict,
        context_window: int = CONTEXT_WINDOW,
    ) -> List[Dict]:
        text: str = para.get("text", "")
        if not text:
            return []
        page_code = para.get("page_code", "")
        full_header = para.get("full_header", "")
        para_idx = para.get("index", 0) + 1

        matches = self._trie.extract(text)
        if not matches:
            return []

        issues: List[Dict] = []
        covered: List[Tuple[int, int]] = []

        for payload, start, end in matches:
            wrong, correct, reason, rule_type = payload

            # Guard 1: ถ้าคำถูกยาวกว่าและข้อความตำแหน่งนั้นเป็นคำถูกอยู่แล้ว ให้ข้าม
            if len(correct) > len(wrong) and text[start:start + len(correct)] == correct:
                continue

            # Guard 2 (Context Awareness — อ่านบริบทคำหน้าประโยค ป้องกัน False Positive):
            prefix_window = text[max(0, start - 16):start]
            suffix_window = text[end:min(len(text), end + 16)]

            # 2.1 เซ็นต์: ห้ามตรวจชน 'เปอร์เซ็นต์' หรือ 'เปอร์ เซ็นต์' เด็ดขาด
            if wrong in ("เซ็นต์", "เซ็นต์ชื่อ", "ลายเซ็นต์", "เซ็น"):
                if "เปอร์" in prefix_window:
                    continue

            # 2.2 องประชุม: แยกแยะระหว่าง 'ห้องประชุม' (สถานที่) และ 'องค์ประชุม' (จำนวนสมาชิก)
            if wrong == "องประชุม":
                clean_prefix = prefix_window.rstrip()
                # หากมีคำว่า 'ห้' อยู่ข้างหน้า (เช่น 'ห้ องประชุม') หรือบริบทห้องประชุม
                if clean_prefix.endswith("ห้") or clean_prefix.endswith("ห้อง") or any(w in prefix_window for w in ("ใน", "เข้า", "ออก", "ติดภารกิจ")):
                    if clean_prefix.endswith("ห้"):
                        h_idx = prefix_window.rfind("ห้")
                        actual_start = max(0, start - 16) + h_idx
                        wrong = text[actual_start:end]
                        correct = "ห้องประชุม"
                        reason = "คำว่า 'ห้ องประชุม' มีการเว้นวรรคหรือตัดคำผิดพลาด ที่ถูกต้องคือ 'ห้องประชุม'"
                        start = actual_start
                    else:
                        # กล่าวถึงห้องประชุมอย่างถูกต้องอยู่แล้ว ให้ข้าม
                        continue
                elif not any(w in prefix_window for w in ("ครบ", "นับ", "ตรวจ", "เป็น", "เสีย", "เปิด")):
                    # หากไม่มีคำบ่งชี้องค์ประชุม (ครบ/นับ/ตรวจ) เลย ให้ข้ามเพื่อไม่ให้เกิด False Positive
                    continue

            # 2.3 เกมส์: อนุญาตชื่อเฉพาะของการแข่งขัน เช่น 'ซีเกมส์', 'โอลิมปิกเกมส์', 'เอเชียนเกมส์'
            if wrong == "เกมส์":
                if any(w in prefix_window for w in ("ซี", "เอเชียน", "โอลิมปิก", "พาราลิมปิก", "เวิลด์", "อาเซียน")):
                    continue

            # 2.4 ชาร์ต: อนุญาตคำศัพท์แผนภูมิ เช่น 'พายชาร์ต', 'โฟลว์ชาร์ต', 'ฟลิปชาร์ต'
            if wrong == "ชาร์ต":
                if any(w in prefix_window for w in ("พาย", "โฟลว์", "ฟลิป")) or any(w in suffix_window for w in ("รูป", "แท่ง", "ข้อมูล")):
                    continue

            # 2.5 โพส: ข้ามถ้าเป็นคำว่า โพสต์ อยู่แล้ว
            if wrong == "โพส" and (text[end:end+2] == "ต์" or text[end:end+3] == "ทรอ"):
                continue

            if any(cs <= start and end <= ce for cs, ce in covered):
                continue
            covered.append((start, end))
            issues.append({
                "rule": rule_type,
                "rule_label": _rule_label(rule_type),
                "para_index": para_idx,
                "page_hint": page_code,
                "page_code": page_code,
                "full_header": full_header,
                "wrong_word": wrong,
                "correct_word": correct,
                "reason": reason,
                "snippet": _build_snippet(text, start, end, context_window),
                "color": _rule_color(rule_type),
            })
        return issues

    def check_paragraphs(self, paragraphs: List[Dict]) -> List[Dict]:
        all_issues: List[Dict] = []
        for para in paragraphs:
            all_issues.extend(self.check_paragraph(para))
        return all_issues

    def total_registered(self) -> int:
        return len(self._registry)


# ─────────────────────────────────────────────────────────────
# 3. RegexChecker
# ─────────────────────────────────────────────────────────────

class RegexChecker:
    """
    ตรวจ Pattern ด้วย Regular Expressions:
    - ไม้ยมก (ๆ): เว้นวรรคหน้า-หลัง
    - เอ.ไอ. -> เอไอ (มติวิปวุฒิสภา)
    - ในชั้นกรรมาธิการ -> ในขั้นกรรมาธิการ
    """

    def __init__(self):
        self._rules: List[Dict] = []
        self._install_builtin_rules()

    def _install_builtin_rules(self) -> None:
        # R1: ไม้ยมก ขาดวรรคหน้า
        self._rules.append({
            "id": "mai_yamok_before",
            "pattern": re.compile(r'(?<=[^\s\(\[\{])ๆ'),
            "handler": self._handle_mai_yamok_before,
            "rule_type": "senate_formatting",
            "label": "ระเบียบสำนักกรรมาธิการ ๓",
        })
        # R2: ไม้ยมก ขาดวรรคหลัง
        self._rules.append({
            "id": "mai_yamok_after",
            "pattern": re.compile(r'ๆ(?=[^\s\)\]\},\.\;\:\?\!\n\r\u0e46])'),
            "handler": self._handle_mai_yamok_after,
            "rule_type": "senate_formatting",
            "label": "ระเบียบสำนักกรรมาธิการ ๓",
        })
        # R3: เอ.ไอ.
        self._rules.append({
            "id": "abbrev_AI",
            "pattern": re.compile(r'เอ\.ไอ\.?'),
            "wrong_word": "เอ.ไอ.",
            "correct_word": "เอไอ",
            "reason": "ตามมติวิปและคู่มือคำทับศัพท์ทางการวุฒิสภา ให้เขียน 'เอไอ' (ไม่มีจุด) หรือใช้ 'ปัญญาประดิษฐ์' เมื่อใช้ครั้งแรก",
            "rule_type": "parliament_rules",
            "label": "ระเบียบวุฒิสภา",
            "handler": None,
        })
        # R4: ในชั้นกรรมาธิการ
        self._rules.append({
            "id": "nai_chan_kam",
            "pattern": re.compile(r'ในชั้นกรรมาธิการ'),
            "wrong_word": "ในชั้นกรรมาธิการ",
            "correct_word": "ในขั้นกรรมาธิการ",
            "reason": "ตามระเบียบงานสารบรรณวุฒิสภาและระเบียบสำนักกรรมาธิการ ๓ ต้องใช้คำว่า 'ในขั้นกรรมาธิการ' (ห้ามใช้ 'ในชั้น')",
            "rule_type": "parliament_rules",
            "label": "ระเบียบวุฒิสภา",
            "handler": None,
        })
        # R5: ในชั้นของกรรมาธิการ
        self._rules.append({
            "id": "nai_chan_khong_kam",
            "pattern": re.compile(r'ในชั้นของกรรมาธิการ'),
            "wrong_word": "ในชั้นของกรรมาธิการ",
            "correct_word": "ในขั้นของกรรมาธิการ",
            "reason": "ตามระเบียบงานสารบรรณวุฒิสภา ต้องใช้คำว่า 'ในขั้นของกรรมาธิการ'",
            "rule_type": "parliament_rules",
            "label": "ระเบียบวุฒิสภา",
            "handler": None,
        })

    def add_rule(
        self,
        pattern: re.Pattern,
        wrong_word: str,
        correct_word: str,
        reason: str,
        rule_type: str = "senate_formatting",
        label: str = "ระเบียบสำนักกรรมาธิการ ๓",
        rule_id: str = "",
    ) -> None:
        self._rules.append({
            "id": rule_id or f"custom_{len(self._rules)}",
            "pattern": pattern,
            "wrong_word": wrong_word,
            "correct_word": correct_word,
            "reason": reason,
            "rule_type": rule_type,
            "label": label,
            "handler": None,
        })

    def _handle_mai_yamok_before(
        self, m: re.Match, text: str, rule: Dict
    ) -> Optional[Dict]:
        idx = m.start()
        start = idx
        for i in range(idx - 1, max(idx - 8, -1), -1):
            if text[i] in ' \t\n\r([{\'\".,;:!?':
                start = i + 1
                break
            start = i
        pre_word = text[start:idx]
        if not pre_word:
            return None
        return {
            "wrong_word": pre_word + "ๆ",
            "correct_word": pre_word + " ๆ",
            "span": (start, idx + 1),
            "reason": "เครื่องหมายไม้ยมก (ๆ) ต้องเว้นวรรคข้างหน้า ๑ เคาะ ตามระเบียบงานสารบรรณ/สำนักกรรมาธิการ ๓",
        }

    def _handle_mai_yamok_after(
        self, m: re.Match, text: str, rule: Dict
    ) -> Optional[Dict]:
        idx = m.start()
        end = idx + 1
        for i in range(idx + 1, min(idx + 8, len(text))):
            if text[i] in ' \t\n\r([{\'\".,;:!?':
                end = i
                break
            end = i + 1
        post_word = text[idx + 1:end]
        if not post_word:
            return None
        return {
            "wrong_word": "ๆ" + post_word,
            "correct_word": "ๆ " + post_word,
            "span": (idx, end),
            "reason": "เครื่องหมายไม้ยมก (ๆ) ต้องเว้นวรรคข้างหลัง ๑ เคาะ ตามระเบียบงานสารบรรณ/สำนักกรรมาธิการ ๓",
        }

    def check_paragraph(
        self, para: Dict, context_window: int = CONTEXT_WINDOW
    ) -> List[Dict]:
        text: str = para.get("text", "")
        if not text:
            return []
        page_code = para.get("page_code", "")
        full_header = para.get("full_header", "")
        para_idx = para.get("index", 0) + 1

        issues: List[Dict] = []
        covered: List[Tuple[int, int]] = []

        for rule in self._rules:
            handler = rule.get("handler")
            rule_type = rule["rule_type"]
            label = rule["label"]

            for m in rule["pattern"].finditer(text):
                start, end = m.start(), m.end()
                if handler:
                    result = handler(m, text, rule)
                    if result is None:
                        continue
                    wrong = result["wrong_word"]
                    correct = result["correct_word"]
                    reason = result["reason"]
                    start, end = result["span"]
                else:
                    wrong = rule.get("wrong_word", m.group())
                    correct = rule.get("correct_word", "")
                    reason = rule.get("reason", "")

                if any(cs <= start and end <= ce for cs, ce in covered):
                    continue
                covered.append((start, end))
                issues.append({
                    "rule": rule_type,
                    "rule_label": label,
                    "para_index": para_idx,
                    "page_hint": page_code,
                    "page_code": page_code,
                    "full_header": full_header,
                    "wrong_word": wrong,
                    "correct_word": correct,
                    "reason": reason,
                    "snippet": _build_snippet(text, start, end, context_window),
                    "color": _rule_color(rule_type),
                })
        return issues

    def check_paragraphs(self, paragraphs: List[Dict]) -> List[Dict]:
        all_issues: List[Dict] = []
        for para in paragraphs:
            all_issues.extend(self.check_paragraph(para))
        return all_issues


# ─────────────────────────────────────────────────────────────
# 4. SenateDocumentChecker — Orchestrator
# ─────────────────────────────────────────────────────────────

class SenateDocumentChecker:
    """รวม DictionaryChecker + RegexChecker — entry point หลัก"""

    def __init__(self, auto_load: bool = True):
        self.dict_checker = DictionaryChecker()
        self.regex_checker = RegexChecker()
        if auto_load:
            self._load_all_databases()

    def _load_all_databases(self) -> None:
        # Priority 1: COMMON_MISPELLING_MAP จาก database_manager (ไม่ overwrite)
        try:
            import database_manager as dm
            loaded = self.dict_checker.load_dict(
                {k: (v, _build_misspelling_reason(k, v)) for k, v in dm.COMMON_MISPELLING_MAP.items()},
                rule_type="misspelling",
                overwrite=False,
            )
            logger.info(f"DictionaryChecker: COMMON_MISPELLING_MAP {loaded} คำ")
        except Exception as e:
            logger.warning(f"ไม่สามารถโหลด COMMON_MISPELLING_MAP: {e}")

        # Priority 2: THAI_TRANSLIT_WRONG_MAP (คู่มือกิจวุฒิ — overwrite สูงสุด)
        loaded2 = self.dict_checker.load_dict(
            THAI_TRANSLIT_WRONG_MAP,
            rule_type="transliteration",
            overwrite=True,
        )
        logger.info(f"DictionaryChecker: THAI_TRANSLIT_WRONG_MAP {loaded2} คำ")

        # Priority 3: EXTRA_TYPO_MAP (ไม่ overwrite)
        loaded3 = self.dict_checker.load_dict(
            EXTRA_TYPO_MAP,
            rule_type="misspelling",
            overwrite=False,
        )
        logger.info(f"DictionaryChecker: EXTRA_TYPO_MAP {loaded3} คำ")
        logger.info(f"DictionaryChecker: รวมทั้งหมด {self.dict_checker.total_registered()} คำ")

    def check(self, paragraphs: List[Dict]) -> List[Dict]:
        """ตรวจ paragraphs ทั้งหมด คืน List[Dict] ผลลัพธ์"""
        all_issues: List[Dict] = []
        all_issues.extend(self.dict_checker.check_paragraphs(paragraphs))
        all_issues.extend(self.regex_checker.check_paragraphs(paragraphs))

        # Deduplicate
        seen: set = set()
        unique: List[Dict] = []
        for iss in all_issues:
            key = (
                iss["para_index"],
                iss["wrong_word"].strip(),
                iss["correct_word"].strip(),
                iss.get("snippet", "").strip(),
            )
            if key not in seen:
                seen.add(key)
                unique.append(iss)

        unique.sort(key=lambda x: (x["para_index"], x.get("rule", "")))
        return unique


# ─────────────────────────────────────────────────────────────
# Helper Functions
# ─────────────────────────────────────────────────────────────

def _build_snippet(text: str, start: int, end: int, window: int = CONTEXT_WINDOW) -> str:
    start_pos = max(0, start - window)
    end_pos = min(len(text), end + window)
    left = ("..." if start_pos > 0 else "") + text[start_pos:start].lstrip()
    wrong = text[start:end]
    right = text[end:end_pos].rstrip() + ("..." if end_pos < len(text) else "")
    return left + wrong + right


def _build_misspelling_reason(wrong: str, correct: str) -> str:
    return (
        f"คำว่า '{wrong}' สะกดผิด ที่ถูกต้องตามฐานข้อมูลคือ '{correct}' "
        "— กรุณาตรวจสอบบริบทก่อนแก้ไข"
    )


def _rule_color(rule_type: str) -> str:
    return {
        "misspelling": "#FF4444",
        "transliteration": "#1ABC9C",
        "parliament_rules": "#E74C3C",
        "senate_formatting": "#E67E22",
        "vocabulary": "#FF8800",
    }.get(rule_type, "#999999")


def _rule_label(rule_type: str) -> str:
    return {
        "misspelling": "คำผิด (ฐานข้อมูล)",
        "transliteration": "คำทับศัพท์",
        "parliament_rules": "ระเบียบวุฒิสภา",
        "senate_formatting": "ระเบียบสำนักกรรมาธิการ ๓",
        "vocabulary": "ศัพท์บัญญัติ",
    }.get(rule_type, rule_type)


# ─────────────────────────────────────────────────────────────
# THAI_TRANSLIT_WRONG_MAP
# อ้างอิง: คู่มือคำทับศัพท์ทางการวุฒิสภา (กิจวุฒิ) — ยึดฐานนี้ก่อนเสมอ
# ─────────────────────────────────────────────────────────────

THAI_TRANSLIT_WRONG_MAP: Dict[str, Tuple[str, str]] = {
    # active -> แอ็กทีฟ
    "แอคทีฟ": ("แอ็กทีฟ", "คำว่า 'แอคทีฟ' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (active) คือ 'แอ็กทีฟ'"),
    "แอ็คทีฟ": ("แอ็กทีฟ", "คำว่า 'แอ็คทีฟ' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (active) คือ 'แอ็กทีฟ'"),
    # application -> แอปพลิเคชัน
    "แอพพลิเคชั่น": ("แอปพลิเคชัน", "คำว่า 'แอพพลิเคชั่น' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (application) คือ 'แอปพลิเคชัน'"),
    "แอพพลิเคชัน": ("แอปพลิเคชัน", "คำว่า 'แอพพลิเคชัน' สะกดผิด ที่ถูกต้องคือ 'แอปพลิเคชัน' (พ ตัวเดียว)"),
    "แอปพลิเคชั่น": ("แอปพลิเคชัน", "คำว่า 'แอปพลิเคชั่น' มีวรรณยุกต์เกิน ที่ถูกต้องคือ 'แอปพลิเคชัน'"),
    "แอพ": ("แอป", "คำว่า 'แอพ' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา คือ 'แอป'"),
    # algorithm -> แอลกอริทึม
    "แอลกอริธึ่ม": ("แอลกอริทึม", "คำว่า 'แอลกอริธึ่ม' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (algorithm) คือ 'แอลกอริทึม'"),
    "แอลกอรึทึม": ("แอลกอริทึม", "คำว่า 'แอลกอรึทึม' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (algorithm) คือ 'แอลกอริทึม'"),
    "อัลกอริธึม": ("แอลกอริทึม", "คำว่า 'อัลกอริธึม' สะกดผิด ที่ถูกต้องคือ 'แอลกอริทึม'"),
    "อัลกอริธึ่ม": ("แอลกอริทึม", "คำว่า 'อัลกอริธึ่ม' สะกดผิด ที่ถูกต้องคือ 'แอลกอริทึม'"),
    "อัลกอริทึ่ม": ("แอลกอริทึม", "คำว่า 'อัลกอริทึ่ม' สะกดผิด ที่ถูกต้องคือ 'แอลกอริทึม'"),
    # absolute -> แอบโซลูต
    "แอบโซลูท": ("แอบโซลูต", "คำว่า 'แอบโซลูท' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (absolute) คือ 'แอบโซลูต' (ใช้ ต เต่า)"),
    # agency -> เอเจนซี
    "เอเจนซี่": ("เอเจนซี", "คำว่า 'เอเจนซี่' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (agency) คือ 'เอเจนซี' (ไม่มีวรรณยุกต์)"),
    # account -> แอ็กเคานต์
    "แอคเคาท์": ("แอ็กเคานต์", "คำว่า 'แอคเคาท์' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (account) คือ 'แอ็กเคานต์'"),
    "แอคเค้าท์": ("แอ็กเคานต์", "คำว่า 'แอคเค้าท์' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา คือ 'แอ็กเคานต์'"),
    "แอ็คเคาท์": ("แอ็กเคานต์", "คำว่า 'แอ็คเคาท์' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา คือ 'แอ็กเคานต์'"),
    "แอกเคาท์": ("แอ็กเคานต์", "คำว่า 'แอกเคาท์' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา คือ 'แอ็กเคานต์'"),
    # big data -> บิ๊กเดตา (คู่มือกิจวุฒิ)
    "บิ๊กดาต้า": ("บิ๊กเดตา", "คำว่า 'บิ๊กดาต้า' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (big data) คือ 'บิ๊กเดตา'"),
    "บิ๊กเดต้า": ("บิ๊กเดตา", "คำว่า 'บิ๊กเดต้า' มีวรรณยุกต์เกิน ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา คือ 'บิ๊กเดตา'"),
    "บิ๊กดาตา": ("บิ๊กเดตา", "คำว่า 'บิ๊กดาตา' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (big data) คือ 'บิ๊กเดตา'"),
    # backlog -> แบ็กล็อก (คู่มือกิจวุฒิ)
    "แบคล็อก": ("แบ็กล็อก", "คำว่า 'แบคล็อก' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (backlog) คือ 'แบ็กล็อก'"),
    "แบ็คล็อก": ("แบ็กล็อก", "คำว่า 'แบ็คล็อก' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (backlog) คือ 'แบ็กล็อก'"),
    # Android -> แอนดรอยด์
    "แอนดรอย": ("แอนดรอยด์", "คำว่า 'แอนดรอย' พิมพ์ตก ด์ ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (Android) คือ 'แอนดรอยด์'"),
    # battery -> แบตเตอรี่ (คู่มือกิจวุฒิ+พจนานุกรม)
    "แบตเตอรี": ("แบตเตอรี่", "คำว่า 'แบตเตอรี' พิมพ์ตกวรรณยุกต์ ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (battery) คือ 'แบตเตอรี่'"),
    # best practice -> เบสต์แพร็กทิซ (คู่มือกิจวุฒิ ระบุว่านี้คือรูปถูกต้อง)
    "เบสต์แพรกทิส": ("เบสต์แพร็กทิซ", "คำว่า 'เบสต์แพรกทิส' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (best practice) คือ 'เบสต์แพร็กทิซ'"),
    "เบสแพรกทิส": ("เบสต์แพร็กทิซ", "คำว่า 'เบสแพรกทิส' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา (best practice) คือ 'เบสต์แพร็กทิซ'"),
    # AI dot variants -> handled by RegexChecker (เอ.ไอ.)
    # smart -> สมาร์ต
    "สมาร์ทโฟน": ("สมาร์ตโฟน", "คำว่า 'สมาร์ทโฟน' สะกดผิด ที่ถูกต้องตามคู่มือคำทับศัพท์ทางการวุฒิสภา คือ 'สมาร์ตโฟน'"),
    "สมาร์ท": ("สมาร์ต", "คำว่า 'สมาร์ท' สะกดผิด ที่ถูกต้องคือ 'สมาร์ต' (ใช้ ต เต่า ไม่มีทัณฑฆาต)"),
    # digital -> ดิจิทัล
    "ดิจิตอล": ("ดิจิทัล", "คำว่า 'ดิจิตอล' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'ดิจิทัล'"),
    # internet -> อินเทอร์เน็ต
    "อินเตอร์เน็ต": ("อินเทอร์เน็ต", "คำว่า 'อินเตอร์เน็ต' สะกดผิดตามราชบัณฑิตยสภา ที่ถูกต้องคือ 'อินเทอร์เน็ต' (ใช้ ท ทหาร)"),
}


# ─────────────────────────────────────────────────────────────
# EXTRA_TYPO_MAP — คำพิมพ์ตกหล่นเพิ่มเติม
# ─────────────────────────────────────────────────────────────

EXTRA_TYPO_MAP: Dict[str, Tuple[str, str]] = {
    "พิจรณา": ("พิจารณา", "คำว่า 'พิจรณา' พิมพ์ตก า ที่ถูกต้องคือ 'พิจารณา'"),
    "ในชั้นกรรมาธิการ": ("ในขั้นกรรมาธิการ", "ตามระเบียบงานสารบรรณวุฒิสภาและระเบียบสำนักกรรมาธิการ ๓ ต้องใช้ 'ในขั้นกรรมาธิการ'"),
    "ในชั้นของกรรมาธิการ": ("ในขั้นของกรรมาธิการ", "ตามระเบียบงานสารบรรณวุฒิสภา ต้องใช้ 'ในขั้นของกรรมาธิการ'"),
    "กรรมาธิกาณ": ("กรรมาธิการ", "คำว่า 'กรรมาธิกาณ' สะกดผิด ที่ถูกต้องคือ 'กรรมาธิการ'"),
    "กรรมาธิการณ": ("กรรมาธิการ", "คำว่า 'กรรมาธิการณ' มี ณ เกินท้าย ที่ถูกต้องคือ 'กรรมาธิการ'"),
    "อนุกรมาธิการ": ("อนุกรรมาธิการ", "คำว่า 'อนุกรมาธิการ' พิมพ์ตก รร ที่ถูกต้องคือ 'อนุกรรมาธิการ'"),
    "อนุกรรมาธิกาณ": ("อนุกรรมาธิการ", "คำว่า 'อนุกรรมาธิกาณ' สะกดผิด ที่ถูกต้องคือ 'อนุกรรมาธิการ'"),
    "ห้ องประชุม": ("ห้องประชุม", "คำว่า 'ห้ องประชุม' มีการเว้นวรรคหรือตัดคำผิดพลาด ที่ถูกต้องคือ 'ห้องประชุม'"),
    "ในห้ องประชุม": ("ในห้องประชุม", "คำว่า 'ในห้ องประชุม' มีการเว้นวรรคหรือตัดคำผิดพลาด ที่ถูกต้องคือ 'ในห้องประชุม'"),
    "เปอร์ เซ็นต์": ("เปอร์เซ็นต์", "คำว่า 'เปอร์เซ็นต์' ให้เขียนติดกัน ไม่ต้องเว้นวรรคกลางคำ ตามคู่มือคำทับศัพท์ทางการวุฒิสภา"),
    "ครบองประชุม": ("ครบองค์ประชุม", "คำว่า 'ครบองประชุม' พิมพ์ตก ค์ ที่ถูกต้องคือ 'ครบองค์ประชุม'"),
    "ไม่ครบองประชุม": ("ไม่ครบองค์ประชุม", "คำว่า 'ไม่ครบองประชุม' พิมพ์ตก ค์ ที่ถูกต้องคือ 'ไม่ครบองค์ประชุม'"),
    "นับองประชุม": ("นับองค์ประชุม", "คำว่า 'นับองประชุม' พิมพ์ตก ค์ ที่ถูกต้องคือ 'นับองค์ประชุม'"),
    "ตรวจองประชุม": ("ตรวจองค์ประชุม", "คำว่า 'ตรวจองประชุม' พิมพ์ตก ค์ ที่ถูกต้องคือ 'ตรวจองค์ประชุม'"),
    "เป็นองประชุม": ("เป็นองค์ประชุม", "คำว่า 'เป็นองประชุม' พิมพ์ตก ค์ ที่ถูกต้องคือ 'เป็นองค์ประชุม'"),
    "มติที่ประขุม": ("มติที่ประชุม", "คำว่า 'มติที่ประขุม' พิมพ์ผิด ข ไข่ ที่ถูกต้องคือ 'มติที่ประชุม'"),
    "ประทานวุฒิสภา": ("ประธานวุฒิสภา", "คำว่า 'ประทานวุฒิสภา' สะกดผิด ที่ถูกต้องคือ 'ประธานวุฒิสภา' (ใช้ ธ ธง)"),
    "รองประทาน": ("รองประธาน", "คำว่า 'รองประทาน' สะกดผิด ที่ถูกต้องคือ 'รองประธาน' (ใช้ ธ ธง)"),
    "พิจารนา": ("พิจารณา", "คำว่า 'พิจารนา' สะกดผิดด้วย น หนู ที่ถูกต้องคือ 'พิจารณา' (ใช้ ณ เณร)"),
    "ผ้อำนวยการ": ("ผู้อำนวยการ", "คำว่า 'ผ้อำนวยการ' พิมพ์ตก ู ที่ถูกต้องคือ 'ผู้อำนวยการ'"),
    "เลขาธิการณ": ("เลขาธิการ", "คำว่า 'เลขาธิการณ' มี ณ เกินท้าย ที่ถูกต้องคือ 'เลขาธิการ'"),
    "นวัฒกรรม": ("นวัตกรรม", "คำว่า 'นวัฒกรรม' สะกดผิดด้วย ฒ ผู้เฒ่า ที่ถูกต้องคือ 'นวัตกรรม' (ใช้ ต เต่า)"),
    "อภิบาย": ("อภิปราย", "คำว่า 'อภิบาย' สะกดผิด ที่ถูกต้องคือ 'อภิปราย'"),
    "สัมฤทธิ์ผล": ("สัมฤทธิผล", "คำว่า 'สัมฤทธิ์ผล' ใส่ทัณฑฆาตที่ ธิ เกินมา ที่ถูกต้องคือ 'สัมฤทธิผล'"),
    "เอกฉันทร์": ("เอกฉันท์", "คำว่า 'เอกฉันทร์' สะกดผิด ที่ถูกต้องคือ 'เอกฉันท์'"),
}


# ─────────────────────────────────────────────────────────────
# Unit Tests
# ─────────────────────────────────────────────────────────────

def _run_unit_tests() -> bool:
    import json
    import sys
    sys.stdout.reconfigure(encoding="utf-8")

    print("=" * 70)
    print("UNIT TEST — SenateDocumentChecker (100% Deterministic)")
    print("=" * 70)

    checker = SenateDocumentChecker(auto_load=True)
    total_reg = checker.dict_checker.total_registered()
    print(f"✓ Trie: ลงทะเบียนคำผิดทั้งหมด {total_reg} คำ\n")

    test_paragraphs = [
        {
            "index": 0,
            "text": (
                "คณะทำงานมีความแอคทีฟในการพัฒนาแอพพลิเคชั่น "
                "และนำแอลกอริธึ่มรวมถึงแอลกอรึทึมมาวิเคราะห์บิ๊กดาต้า "
                "พร้อมทั้งยืนยันด้วยระบบแอบโซลูทของเอเจนซี่"
            ),
            "page_code": "จันทร์ตรี ๑/๑",
            "full_header": "จันทร์ตรี ๑/๑",
        },
        {
            "index": 1,
            "text": (
                "ผู้ใช้งานเชื่อมต่อแอคเคาท์จากแอนดรอยและชาร์จแบตเตอรี "
                "พร้อมทั้งเคลียร์แบคล็อกตามเบสต์แพรกทิสที่กำหนด"
            ),
            "page_code": "จันทร์ตรี ๑/๒",
            "full_header": "จันทร์ตรี ๑/๒",
        },
        {
            "index": 2,
            "text": "ที่ประชุมได้มีมติพิจรณาเรื่องในชั้นกรรมาธิการด้วยระบบเอ.ไอ.",
            "page_code": "จันทร์ตรี ๑/๓",
            "full_header": "จันทร์ตรี ๑/๓",
        },
        {
            "index": 3,
            "text": "มีปัญหาต่างๆนานา ที่ต้องหารือในชั้นกรรมาธิการอีกต่างๆ",
            "page_code": "จันทร์ตรี ๒/๑",
            "full_header": "จันทร์ตรี ๒/๑",
        },
        {
            "index": 4,
            "text": "ผลการดำเนินงาน ต่าง ๆ เป็นที่น่าพอใจ",  # ถูกต้องแล้ว
            "page_code": "จันทร์ตรี ๒/๒",
            "full_header": "จันทร์ตรี ๒/๒",
        },
    ]

    # คำที่ต้องตรวจเจอทั้งหมด (ขั้นต่ำ)
    EXPECTED_WORDS = {
        "แอคทีฟ", "แอพพลิเคชั่น",
        "แอลกอริธึ่ม", "แอลกอรึทึม",
        "บิ๊กดาต้า", "แอบโซลูท", "เอเจนซี่",
        "แอคเคาท์", "แบตเตอรี", "แบคล็อก", "เบสต์แพรกทิส",
        "พิจรณา",
        "ในชั้นกรรมาธิการ",
        "เอ.ไอ.",
    }

    issues = checker.check(test_paragraphs)
    found = {iss["wrong_word"] for iss in issues}
    missed = EXPECTED_WORDS - found

    print(f"ตรวจพบทั้งหมด {len(issues)} รายการ\n")
    for iss in issues:
        print(
            f"  [ย่อหน้า {iss['para_index']}] [{iss['rule_label']}] "
            f"'{iss['wrong_word']}' -> '{iss['correct_word']}'"
        )
        print(f"    บริบท: ...{iss['snippet']}...")
    print()

    print("─" * 70)
    if missed:
        print(f"❌ FAIL — ตรวจหลุด {len(missed)} คำ: {missed}")
        return False
    else:
        print(f"✅ PASS — ตรวจจับครบ 100% ทุก {len(EXPECTED_WORDS)} คำทดสอบ ไม่มีคำหลุดเลย!")
        return True


if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    success = _run_unit_tests()
    exit(0 if success else 1)
