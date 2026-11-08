# -*- coding: utf-8 -*-
"""
Excel 模板生成模块（动态识别模板布局）。

保存路径：VolunteerStatistics/core/excel_exporter.py

功能：
  - 依据真实模板，按班级生成独立 Excel 文件
  - 自动填写学生志愿服务信息（学号、姓名、项目次数、总时长、各项目时长）
  - 自动统计班级总时长与每月人均
  - 动态生成项目列（超过模板默认 4 列时自动插入新列）
  - 尽可能完整保留原始模板样式（直接复制模板原地填值）
  - 不修改原始模板
  - 支持多班级批量导出、打包为 ZIP

支持两种模板布局（视传2501 / 视传2502），通过动态识别以下位置：
  - 汇总标签「班级总时长」所在合并单元格（据此确定汇总值列）
  - 底部「带队干事」行（据此确定数据区行数）
  - 数据行的行级合并单元格（如 2502 的 I:K 逐行合并）
"""
import zipfile
from collections import Counter, OrderedDict
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

# 基础布局常量（两种模板一致的部分）
PROJECT_START_COL = 5   # E 项目起始列
PROJECT_BASE_COUNT = 4  # 模板默认 4 个项目列（E~H）
DATA_START_ROW = 4      # 学生数据起始行
PROJECT_COL_WIDTH = 16.5  # 新增项目列宽度


# ---------------------------------------------------------------------------
# 样式复制
# ---------------------------------------------------------------------------
def _copy_cell_style(src, dst):
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def _copy_row_style(ws, src_row, dst_row):
    for c in range(1, ws.max_column + 1):
        _copy_cell_style(
            ws.cell(row=src_row, column=c),
            ws.cell(row=dst_row, column=c),
        )
    if src_row in ws.row_dimensions and ws.row_dimensions[src_row].height:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


# ---------------------------------------------------------------------------
# 布局检测
# ---------------------------------------------------------------------------
def _detect_layout(ws):
    """动态检测模板布局，返回汇总列、底部行等关键位置。"""
    summary_min = summary_max = None
    for mc in ws.merged_cells.ranges:
        v = ws.cell(row=2, column=mc.min_col).value
        if v and "班级总时长" in str(v):
            summary_min, summary_max = mc.min_col, mc.max_col
            break
    if summary_min is None:
        for c in range(PROJECT_START_COL, ws.max_column + 1):
            v = ws.cell(row=2, column=c).value
            if v and "班级总时长" in str(v):
                summary_min = summary_max = c
                break

    footer_row = ws.max_row
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and "带队干事" in str(v):
            footer_row = r
            break

    return {
        "summary_min": summary_min,
        "summary_max": summary_max,
        "summary_value_col": summary_max + 1 if summary_max else 15,
        "footer_row": footer_row,
        "data_start": DATA_START_ROW,
    }


def _detect_data_row_merge_cols(ws, data_start, footer_row):
    """检测数据行的行级合并列区间（除 A:B 外），返回 (min_col, max_col) 或 None。"""
    spans = []
    for mc in ws.merged_cells.ranges:
        if mc.min_col == 1 and mc.max_col == 2:
            continue
        if mc.min_row >= data_start and mc.max_row < footer_row:
            spans.append((mc.min_col, mc.max_col))
    if spans:
        return Counter(spans).most_common(1)[0][0]
    return None


# ---------------------------------------------------------------------------
# 列插入（项目数超过 4 时）
# ---------------------------------------------------------------------------
def _insert_cols_before(ws, insert_at_col, amount):
    """在 insert_at_col 前插入 amount 列，并修正合并单元格与列宽。"""
    affected = []
    for mc in list(ws.merged_cells.ranges):
        if mc.min_col >= insert_at_col:
            affected.append((mc.min_row, mc.min_col, mc.max_row, mc.max_col))
            ws.merged_cells.ranges.remove(mc)

    old_widths = {}
    for letter, dim in list(ws.column_dimensions.items()):
        if dim.width is not None:
            old_widths[column_index_from_string(letter)] = dim.width
    for letter in list(ws.column_dimensions.keys()):
        del ws.column_dimensions[letter]

    ws.insert_cols(insert_at_col, amount)

    for (r1, c1, r2, c2) in affected:
        ws.merge_cells(
            start_row=r1, start_column=c1 + amount,
            end_row=r2, end_column=c2 + amount,
        )
    for col_idx, w in old_widths.items():
        new_idx = col_idx + amount if col_idx >= insert_at_col else col_idx
        ws.column_dimensions[get_column_letter(new_idx)].width = w
    for i in range(amount):
        ws.column_dimensions[get_column_letter(insert_at_col + i)].width = PROJECT_COL_WIDTH


# ---------------------------------------------------------------------------
# 数据行数量调整
# ---------------------------------------------------------------------------
def _fit_data_rows(ws, data_start, footer_row, student_count):
    """调整数据区行数，并正确管理行级合并与「带队干事」合并。"""
    merge_cols = _detect_data_row_merge_cols(ws, data_start, footer_row)
    slots = footer_row - data_start

    # 1. 移除数据区行级合并（非 A:B）
    for mc in list(ws.merged_cells.ranges):
        if mc.min_col == 1 and mc.max_col == 2:
            continue
        if mc.min_row >= data_start and mc.max_row < footer_row:
            ws.merged_cells.ranges.remove(mc)

    # 2. 移除 footer 合并（A:B 且行 >= data_start）
    for mc in list(ws.merged_cells.ranges):
        if mc.min_col == 1 and mc.max_col == 2 and mc.min_row >= data_start:
            ws.merged_cells.ranges.remove(mc)

    # 3. 插入/删除数据行
    if student_count > slots:
        diff = student_count - slots
        ws.insert_rows(footer_row, diff)
        for i in range(diff):
            _copy_row_style(ws, data_start, footer_row + i)
    elif student_count < slots:
        diff = slots - student_count
        ws.delete_rows(data_start + student_count, diff)

    # 4. 重建数据区行级合并
    if merge_cols:
        for r in range(data_start, data_start + student_count):
            ws.merge_cells(
                start_row=r, start_column=merge_cols[0],
                end_row=r, end_column=merge_cols[1],
            )

    # 5. 重建 footer 合并
    new_footer = data_start + student_count
    ws.merge_cells(start_row=new_footer, start_column=1, end_row=new_footer, end_column=2)


# ---------------------------------------------------------------------------
# 数据准备
# ---------------------------------------------------------------------------
def build_class_data(roster_students, per_student_entries):
    """合并名单与统计结果，生成导出数据。

    返回 (students, projects)：
      students: [{student_id, name, count, total_hours, project_hours}, ...]
      projects: 有序项目名列表
    """
    stats_by_key = {}
    for e in per_student_entries:
        s = e["student"]
        key = (
            s.get("class_name_normalized"),
            s.get("student_id_normalized") or s.get("name_normalized"),
        )
        stats_by_key[key] = e

    project_order = OrderedDict()
    for e in per_student_entries:
        for sr in e["records"]:
            p = str(sr["record"].get("project_name") or "").strip()
            if p and p not in project_order:
                project_order[p] = None
    projects = list(project_order.keys())

    students = []
    for st in roster_students:
        key = (
            st.get("class_name_normalized"),
            st.get("student_id_normalized") or st.get("name_normalized"),
        )
        e = stats_by_key.get(key)
        if e:
            project_hours = OrderedDict()
            for sr in e["records"]:
                p = str(sr["record"].get("project_name") or "").strip()
                if p and sr["hours"] is not None:
                    project_hours[p] = project_hours.get(p, 0.0) + sr["hours"]
            students.append(
                {
                    "student_id": st["student_id"],
                    "name": st["name"],
                    "count": e["count"],
                    "total_hours": round(e["total_hours"], 2),
                    "project_hours": {p: round(h, 2) for p, h in project_hours.items()},
                }
            )
        else:
            students.append(
                {
                    "student_id": st["student_id"],
                    "name": st["name"],
                    "count": 0,
                    "total_hours": 0.0,
                    "project_hours": {},
                }
            )
    return students, projects


# ---------------------------------------------------------------------------
# 导出单班级
# ---------------------------------------------------------------------------
def export_class(template_path, class_name, students, projects, output_path):
    """按模板生成一个班级的 Excel 文件（不修改原模板）。"""
    wb = load_workbook(template_path)
    ws = wb.active
    layout = _detect_layout(ws)

    p_count = len(projects)
    extra = max(0, p_count - PROJECT_BASE_COUNT)

    # 1. 项目数超过 4 时，在 H 列之后（第 9 列）插入新列
    if extra > 0:
        _insert_cols_before(ws, PROJECT_START_COL + PROJECT_BASE_COUNT, extra)

    # 2. 清空并写入项目表头（自 E 列起，至少 4 个槽位）
    total_slots = max(PROJECT_BASE_COUNT, p_count)
    for i in range(total_slots):
        ws.cell(row=2, column=PROJECT_START_COL + i).value = None
    for i, pname in enumerate(projects):
        ws.cell(row=2, column=PROJECT_START_COL + i).value = pname

    # 3. 标题 = 班级名
    ws.cell(row=1, column=1).value = class_name

    # 4. 调整数据行数量
    _fit_data_rows(ws, layout["data_start"], layout["footer_row"], len(students))

    # 5. 填写学生数据
    for i, stu in enumerate(students):
        r = layout["data_start"] + i
        ws.cell(row=r, column=1).value = stu["student_id"] or None
        ws.cell(row=r, column=2).value = stu["name"] or None
        if stu["count"]:
            ws.cell(row=r, column=3).value = stu["count"]
            ws.cell(row=r, column=4).value = stu["total_hours"]
        else:
            ws.cell(row=r, column=3).value = None
            ws.cell(row=r, column=4).value = None
        for j in range(total_slots):
            pname = projects[j] if j < p_count else None
            val = stu["project_hours"].get(pname) if pname else None
            ws.cell(row=r, column=PROJECT_START_COL + j).value = val

    # 6. 汇总：班级总时长、每月人均
    value_col = layout["summary_value_col"] + extra
    total = round(sum(s["total_hours"] for s in students), 2)
    avg = round(total / len(students), 2) if students else 0.0
    ws.cell(row=2, column=value_col).value = total
    ws.cell(row=3, column=value_col).value = avg

    wb.save(output_path)
    return str(output_path)


# ---------------------------------------------------------------------------
# 批量导出与打包
# ---------------------------------------------------------------------------
def export_classes(template_path, classes, output_dir):
    """批量导出多个班级。classes: [(class_name, students, projects), ...]"""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for class_name, students, projects in classes:
        out = output_dir / f"{class_name}.xlsx"
        export_class(template_path, class_name, students, projects, out)
        paths.append(str(out))
    return paths


def zip_files(paths, zip_path):
    """把多个文件打包为 ZIP。"""
    zip_path = Path(zip_path)
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in paths:
            zf.write(p, arcname=Path(p).name)
    return str(zip_path)
