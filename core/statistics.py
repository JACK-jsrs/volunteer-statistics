# -*- coding: utf-8 -*-
"""
志愿服务统计与审核支撑模块。

保存路径：VolunteerStatistics/core/statistics.py

功能：
  - 服务时长解析（小时/分钟单位转换）
  - 按学生统计服务项目、服务次数、总时长
  - 保留同一学生参加多个项目的记录
  - 疑似重复记录检测（按 学号/姓名 + 项目 + 日期 分组）
  - 缺失时长 / 无法匹配 / 同名 / 疑似重复 分类
  - 支持人工修改（字段覆盖）与排除记录，修改后重新计算
  - 每条记录保留原始文件与原始位置，便于追溯
"""
import re
from collections import OrderedDict

from core.student_matcher import normalize_text


def parse_duration(text):
    """把服务时长文本解析为小时数（浮点）；无法解析返回 None。

    支持：纯数字(按小时)、X小时、X分钟、X小时Y分、H:MM、Xh Ym、小数小时、全角数字等。
    """
    if text is None:
        return None
    s = normalize_text(str(text)).lower()
    if not s:
        return None

    s = s.replace("半小时", "0.5小时")
    s = s.replace("小时", "h").replace("时", "h").replace("分钟", "m").replace("分", "m")

    # Xh Ym 形式
    parts = re.findall(r"([+-]?\d+(?:\.\d+)?)\s*(h|m)", s)
    if parts:
        total = 0.0
        for num, unit in parts:
            v = float(num)
            total += v if unit == "h" else v / 60.0
        return round(total, 4)

    # H:MM 形式
    m = re.fullmatch(r"(\d{1,3}):([0-5]?\d)", s)
    if m:
        return round(int(m.group(1)) + int(m.group(2)) / 60.0, 4)

    # 纯数字（按小时）
    m = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)", s)
    if m:
        return float(m.group(1))

    return None


def format_hours(hours):
    """把小时数格式化为「X小时Y分」。"""
    if hours is None:
        return ""
    total_minutes = round(hours * 60)
    h = total_minutes // 60
    m = total_minutes % 60
    if h and m:
        return f"{h}小时{m}分"
    if h:
        return f"{h}小时"
    return f"{m}分"


def record_id(record):
    """记录唯一标识：原始文件 + 原始位置（用于编辑/排除的键）。"""
    return f"{record.get('source_file') or ''}||{record.get('source_location') or ''}"


def student_unique_key(student):
    cls = student.get("class_name_normalized") or ""
    sid = student.get("student_id_normalized") or ""
    name = student.get("name_normalized") or ""
    return (cls, sid) if sid else (cls, name)


def build_stats_records(match_results, edits=None, excluded=None):
    """把匹配结果转为统计记录，应用人工修改与排除。"""
    edits = edits or {}
    excluded = excluded or {}
    stats_records = []
    for r in match_results:
        if r["status"] != "matched":
            continue
        student = r["matched_student"]
        rec = dict(r["record"])
        rid = record_id(rec)
        if rid in excluded:
            continue
        if rid in edits:
            for k, v in edits[rid].items():
                if v not in (None, ""):
                    rec[k] = v
        stats_records.append(
            {
                "record_id": rid,
                "student": student,
                "record": rec,
                "hours": parse_duration(rec.get("service_hours")),
                "hours_text": rec.get("service_hours", ""),
            }
        )
    return stats_records


def detect_duplicates(stats_records):
    """按 (学号/姓名, 项目, 日期) 分组，返回疑似重复组。"""
    groups = OrderedDict()
    for sr in stats_records:
        sid = (
            sr["student"].get("student_id_normalized")
            or sr["student"].get("name_normalized")
            or ""
        )
        project = normalize_text(sr["record"].get("project_name") or "")
        date = normalize_text(sr["record"].get("service_date") or "")
        key = (sid, project, date)
        groups.setdefault(key, []).append(sr)
    return [
        {"key": key, "records": items}
        for key, items in groups.items()
        if len(items) > 1
    ]


def compute_statistics(stats_records):
    """按学生聚合：项目列表、次数、总时长；返回 (per_student, missing_hours)。"""
    by_student = OrderedDict()
    missing_hours = []
    for sr in stats_records:
        key = student_unique_key(sr["student"])
        entry = by_student.setdefault(
            key,
            {
                "student": sr["student"],
                "records": [],
                "projects": [],
                "count": 0,
                "total_hours": 0.0,
                "hours_missing": 0,
            },
        )
        entry["records"].append(sr)
        entry["count"] += 1
        pname = str(sr["record"].get("project_name") or "").strip()
        if pname and pname not in entry["projects"]:
            entry["projects"].append(pname)
        if sr["hours"] is None:
            entry["hours_missing"] += 1
            missing_hours.append(sr)
        else:
            entry["total_hours"] += sr["hours"]
    return list(by_student.values()), missing_hours


def categorize(match_results, stats_records):
    """分类：缺失时长、无法匹配、同名、疑似重复。"""
    return {
        "missing_hours": [sr for sr in stats_records if sr["hours"] is None],
        "unmatched": [r for r in match_results if r["status"] == "unmatched"],
        "ambiguous": [r for r in match_results if r["status"] == "ambiguous"],
        "duplicates": detect_duplicates(stats_records),
    }
