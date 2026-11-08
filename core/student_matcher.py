# -*- coding: utf-8 -*-
"""
班级学生名单管理与自动匹配模块。

保存路径：VolunteerStatistics/core/student_matcher.py

功能：
  - 加载班级学生名单（Excel，复用 document_parser 读取）
  - 班级与名单关联（手动指定 / 自动识别「班级」列）
  - 学号精确匹配优先；学号缺失时用姓名匹配
  - 姓名规范化：去除前后空格、全角转半角、忽略内部空格
  - 同名学生不自动认定为同一人（进入待人工确认）
  - 无法匹配的学生进入异常列表
  - 支持人工指定正确学生
  - 只统计用户指定的班级
"""
import re

from core import document_parser


# ---------------------------------------------------------------------------
# 文本规范化
# ---------------------------------------------------------------------------
def _fullwidth_to_halfwidth(text):
    """全角字符/全角空格转半角。"""
    result = []
    for ch in str(text):
        code = ord(ch)
        if code == 0x3000:  # 全角空格
            result.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:  # 全角字符
            result.append(chr(code - 0xFEE0))
        else:
            result.append(ch)
    return "".join(result)


def normalize_text(text):
    """全角转半角，并去除所有空白（含前后空格与内部空格）。"""
    if text is None:
        return ""
    s = _fullwidth_to_halfwidth(str(text))
    return re.sub(r"\s+", "", s).strip()


def normalize_name(name):
    return normalize_text(name)


def normalize_class(class_name):
    return normalize_text(class_name)


def normalize_student_id(sid):
    return normalize_text(sid).upper()


def record_signature(record):
    """生成记录的稳定标识，用于人工指定的持久化。"""
    return (
        record.get("source_file", ""),
        record.get("source_location", ""),
        normalize_name(record.get("name", "")),
        normalize_student_id(record.get("student_id", "")),
    )


class StudentMatcher:
    """班级名单管理 + 自动匹配。"""

    def __init__(self, specified_classes=None):
        self.rosters = {}             # norm_class -> {"class_name": str, "students": [...]}
        self.specified_classes = []   # 原始班级名列表
        self._specified_norms = set()
        self._students_by_id = {}
        self._students_by_name = {}
        self.manual = {}              # record_signature -> student dict
        if specified_classes:
            self.set_specified_classes(specified_classes)

    # ------------------------------------------------------------------
    # 班级
    # ------------------------------------------------------------------
    def set_specified_classes(self, classes):
        self.specified_classes = [str(c).strip() for c in classes if str(c).strip()]
        self._specified_norms = {normalize_class(c) for c in self.specified_classes}

    def _specified_students(self):
        students = []
        for norm, entry in self.rosters.items():
            if norm in self._specified_norms:
                students.extend(entry["students"])
        return students

    def _students_in_class(self, norm_class):
        entry = self.rosters.get(norm_class)
        return entry["students"] if entry else []

    # ------------------------------------------------------------------
    # 名单加载
    # ------------------------------------------------------------------
    def load_roster(self, excel_path, class_name=None):
        """从 Excel 加载学生名单。

        class_name 给定则整份名单归入该类；否则从名单中的「班级」列自动识别。
        返回 {"ok", "students_added", "skipped", "errors"}。
        """
        result = document_parser.parse_excel(excel_path)
        if result["errors"]:
            return {
                "ok": False,
                "students_added": 0,
                "skipped": 0,
                "errors": result["errors"],
            }

        added = 0
        skipped = 0
        for rec in result["records"]:
            name = str(rec.get("name") or "").strip()
            sid = str(rec.get("student_id") or "").strip()
            if not name and not sid:
                skipped += 1
                continue
            cls = class_name or rec.get("class_name") or ""
            cls = str(cls).strip()
            if not cls:
                skipped += 1
                continue
            if self.add_student(cls, name, sid, str(excel_path)):
                added += 1
            else:
                skipped += 1
        return {"ok": True, "students_added": added, "skipped": skipped, "errors": []}

    def add_student(self, class_name, name, student_id="", source_file=""):
        """向指定班级添加一名学生；重复返回 False。"""
        norm_class = normalize_class(class_name)
        student = {
            "name": str(name).strip(),
            "name_normalized": normalize_name(name),
            "student_id": str(student_id).strip(),
            "student_id_normalized": normalize_student_id(student_id),
            "class_name": str(class_name).strip(),
            "class_name_normalized": norm_class,
            "source_file": str(source_file),
        }
        if not student["name_normalized"] and not student["student_id_normalized"]:
            return False

        entry = self.rosters.setdefault(
            norm_class, {"class_name": str(class_name).strip(), "students": []}
        )
        for s in entry["students"]:
            if (
                student["student_id_normalized"]
                and s["student_id_normalized"] == student["student_id_normalized"]
            ):
                return False
            if (
                not student["student_id_normalized"]
                and not s["student_id_normalized"]
                and s["name_normalized"] == student["name_normalized"]
            ):
                return False

        entry["students"].append(student)
        self._rebuild_index()
        return True

    def _rebuild_index(self):
        self._students_by_id = {}
        self._students_by_name = {}
        for entry in self.rosters.values():
            for s in entry["students"]:
                if s["student_id_normalized"]:
                    self._students_by_id.setdefault(s["student_id_normalized"], []).append(s)
                if s["name_normalized"]:
                    self._students_by_name.setdefault(s["name_normalized"], []).append(s)

    # ------------------------------------------------------------------
    # 查询
    # ------------------------------------------------------------------
    def get_class_names(self):
        return [entry["class_name"] for entry in self.rosters.values()]

    def get_students(self, class_name=None):
        """返回指定班级，或全部指定班级范围内的学生列表。"""
        if class_name:
            return self._students_in_class(normalize_class(class_name))
        return self._specified_students()

    # ------------------------------------------------------------------
    # 匹配
    # ------------------------------------------------------------------
    def match_records(self, records):
        results = []
        for i, rec in enumerate(records):
            r = self.match_record(rec)
            r["index"] = i
            results.append(r)
        return results

    def match_record(self, record):
        sig = record_signature(record)
        if sig in self.manual:
            s = self.manual[sig]
            return self._result(record, "matched", s, [s], "人工指定", "manual")

        rec_class = normalize_class(record.get("class_name") or "")
        sid = normalize_student_id(record.get("student_id") or "")
        name = normalize_name(record.get("name") or "")

        # 记录带班级但不属于指定班级 → 排除
        if rec_class and rec_class not in self._specified_norms:
            return self._result(record, "excluded", None, [], "班级不在指定统计班级内", None)

        candidates = self._students_in_class(rec_class) if rec_class else self._specified_students()
        if not candidates:
            return self._result(record, "unmatched", None, [], "指定班级名单为空", None)

        # 1) 学号精确匹配（优先）
        if sid:
            id_matches = [s for s in candidates if s["student_id_normalized"] == sid]
            if len(id_matches) == 1:
                s = id_matches[0]
                note = None
                if name and s["name_normalized"] != name:
                    note = f"学号匹配，但姓名不一致（名单为「{s['name']}」）"
                return self._result(record, "matched", s, id_matches, note or "学号精确匹配", "student_id")
            if len(id_matches) > 1:
                return self._result(record, "ambiguous", None, id_matches, "学号在名单中重复", "student_id")
            name_candidates = [s for s in candidates if name and s["name_normalized"] == name]
            return self._result(record, "unmatched", None, name_candidates, "学号未在名单中找到", "student_id")

        # 2) 姓名匹配（学号缺失）
        if name:
            name_matches = [s for s in candidates if s["name_normalized"] == name]
            if len(name_matches) == 1:
                return self._result(record, "matched", name_matches[0], name_matches, "姓名匹配", "name")
            if len(name_matches) > 1:
                return self._result(record, "ambiguous", None, name_matches, "同名学生，需人工确认", "name")
            return self._result(record, "unmatched", None, [], "姓名未在名单中找到", "name")

        return self._result(record, "unmatched", None, [], "记录缺少姓名和学号", None)

    def assign_manually(self, record, student):
        """人工把记录指定为某学生。"""
        self.manual[record_signature(record)] = student

    # ------------------------------------------------------------------
    # 汇总
    # ------------------------------------------------------------------
    def summarize(self, results):
        counts = {"matched": 0, "unmatched": 0, "ambiguous": 0, "excluded": 0}
        for r in results:
            counts[r["status"]] = counts.get(r["status"], 0) + 1
        return counts

    @staticmethod
    def _result(record, status, matched, candidates, reason, match_key):
        return {
            "record": record,
            "status": status,
            "matched_student": matched,
            "candidates": candidates,
            "reason": reason,
            "match_key": match_key,
        }
