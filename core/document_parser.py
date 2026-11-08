# -*- coding: utf-8 -*-
"""
多格式文档解析模块。

保存路径：VolunteerStatistics/core/document_parser.py

功能：
  - Excel（.xlsx）：读取所有工作表、表头、单元格内容
  - Word（.docx）：读取段落与表格
  - PDF：提取文字与表格（依赖 PyMuPDF）
  - 图片（.jpg/.jpeg/.png）：本批次仅识别，不做 OCR
  - 将不同来源的信息尽量转换为统一记录格式
  - 保留原始内容与来源信息，无法解析的文件进入错误列表
  - 不编造缺失信息（缺失字段一律留空）

统一记录字段：
  姓名(name)、学号(student_id)、班级(class_name)、项目名称(project_name)、
  服务日期(service_date)、服务时长(service_hours)、原始文件名(source_file)、
  原始位置(source_location)；另保留 raw（原始单元格/行内容）。
"""
import datetime as _dt
import re
from collections import OrderedDict
from pathlib import Path

import openpyxl
from docx import Document

try:
    import fitz  # PyMuPDF
    _HAS_PYMUPDF = True
except Exception:  # noqa: BLE001
    fitz = None
    _HAS_PYMUPDF = False

# ---------------------------------------------------------------------------
# 统一记录字段定义
# ---------------------------------------------------------------------------
RECORD_FIELDS = [
    "name",
    "student_id",
    "class_name",
    "project_name",
    "service_date",
    "service_hours",
]

FIELD_LABELS = OrderedDict(
    [
        ("name", "姓名"),
        ("student_id", "学号"),
        ("class_name", "班级"),
        ("project_name", "项目名称"),
        ("service_date", "服务日期"),
        ("service_hours", "服务时长"),
    ]
)

SOURCE_FILE_LABEL = "原始文件名"
SOURCE_LOCATION_LABEL = "原始位置"

# 字段别名：用于匹配不同文件中可能出现的各种表头
FIELD_ALIASES = OrderedDict(
    [
        ("name", ["姓名", "名字", "学生姓名", "志愿者姓名", "姓　名"]),
        ("student_id", ["学号", "学生学号", "学籍号", "工号", "编号", "学号/工号"]),
        ("class_name", ["班级", "班级名称", "所属班级", "行政班", "班别"]),
        (
            "project_name",
            ["项目名称", "服务项目", "活动名称", "志愿项目", "服务内容", "志愿活动", "活动", "服务事项", "项目"],
        ),
        ("service_date", ["服务日期", "活动日期", "志愿日期", "服务时间", "活动时间", "志愿时间", "日期"]),
        (
            "service_hours",
            ["服务时长", "志愿时长", "服务小时", "时长", "小时数", "工时", "时长(小时)", "时长（小时）"],
        ),
    ]
)


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def _normalize(text):
    """规范化表头/字段名，便于匹配。"""
    if text is None:
        return ""
    s = str(text)
    for full, half in [("（", "("), ("）", ")"), ("：", ":"), ("　", " ")]:
        s = s.replace(full, half)
    s = re.sub(r"[\s:：()（）*\t]+", "", s)
    return s.lower()


def _build_alias_map():
    mapping = {}
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            mapping.setdefault(_normalize(alias), field)
    return mapping


_ALIAS_MAP = _build_alias_map()


def match_field(header):
    """根据表头匹配统一字段；未匹配返回 None。"""
    h = _normalize(header)
    if not h:
        return None
    # 1) 精确匹配
    if h in _ALIAS_MAP:
        return _ALIAS_MAP[h]
    # 2) 包含匹配（仅对较长别名，且表头长度接近别名，避免标题等长句误判）
    for key, field in _ALIAS_MAP.items():
        if len(key) < 3 or len(h) < 2:
            continue
        if key in h and len(h) <= len(key) + 3:
            return field
    return None


def _to_text(value):
    """把单元格/文本值转换为字符串；空值转为空字符串。"""
    if value is None:
        return ""
    if isinstance(value, _dt.datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0 and value.microsecond == 0:
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, _dt.date):
        return value.isoformat()
    if isinstance(value, _dt.time):
        return value.isoformat()
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (int, float)):
        return str(value)
    return str(value).strip()


def _is_empty_row(row):
    return all(_to_text(c) == "" for c in row)


def _find_header_row(rows, max_scan=10):
    """在前几行中定位表头行：选择可识别字段最多的那一行。"""
    best_idx = 0
    best_count = 0
    for i, row in enumerate(rows[:max_scan]):
        matched = {match_field(_to_text(cell)) for cell in row}
        matched.discard(None)
        if len(matched) > best_count:
            best_count = len(matched)
            best_idx = i
    return best_idx


def _halfwidth(text):
    """全角字符转半角。"""
    result = []
    for ch in str(text):
        code = ord(ch)
        if code == 0x3000:
            result.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        else:
            result.append(ch)
    return "".join(result)


def _looks_like_form_table(rows):
    """判断表格是否为「标签|值」横向表单（如 姓名|胡家维|出生年月|20080328）。"""
    if not rows:
        return False
    row0 = rows[0]
    for i in range(0, len(row0) - 1, 2):
        label = _to_text(row0[i])
        value = _to_text(row0[i + 1])
        if match_field(label) and value and not match_field(value):
            return True
    return False


def _extract_fields_from_text(text):
    """从一段文本中按「字段名：值」等模式提取字段，未识别返回空 dict。

    匹配前先去除所有空白并全角转半角，以兼容「姓  名：张三」这类带空格的标签。
    """
    if not text:
        return {}
    compact = re.sub(r"\s+", "", _halfwidth(text))
    fields = {}
    for field, aliases in FIELD_ALIASES.items():
        if field in fields:
            continue
        for alias in aliases:
            pattern = re.compile(re.escape(_halfwidth(alias)) + r"[:：=]([^,，;；、|\n]+)")
            m = pattern.search(compact)
            if m:
                fields[field] = m.group(1).strip(" 　：:=，,；;")
                break
    return fields


def _new_result(file_path, file_type):
    return {
        "file": str(file_path),
        "file_type": file_type,
        "records": [],
        "raw_texts": [],
        "errors": [],
        "warnings": [],
    }


def _build_record(mapped, source_file, source_location, raw=None):
    """把已匹配字段构造成统一记录；缺失字段留空，不编造。"""
    record = {f: "" for f in RECORD_FIELDS}
    for f in RECORD_FIELDS:
        v = mapped.get(f)
        record[f] = v if v else ""
    record["source_file"] = source_file
    record["source_location"] = source_location
    record["raw"] = raw or {}
    return record


def _rows_to_records(rows, source_file, location_prefix):
    """把「表头 + 数据行」结构转换为统一记录列表。"""
    records = []
    if not rows:
        return records
    header_idx = _find_header_row(rows)
    headers = [_to_text(c) for c in rows[header_idx]]
    for r_idx, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if _is_empty_row(row):
            continue
        raw = {}
        mapped = {}
        for c_idx, val in enumerate(row):
            header = headers[c_idx] if c_idx < len(headers) else f"列{c_idx + 1}"
            if not header:
                header = f"列{c_idx + 1}"
            txt = _to_text(val)
            raw[header] = txt
            field = match_field(header)
            if field and (field not in mapped or (mapped[field] == "" and txt != "")):
                mapped[field] = txt
        records.append(
            _build_record(mapped, source_file, f"{location_prefix}第{r_idx}行", raw)
        )
    return records


# ---------------------------------------------------------------------------
# Excel 解析
# ---------------------------------------------------------------------------
def parse_excel(file_path):
    file_path = Path(file_path)
    result = _new_result(file_path, "Excel")
    if file_path.suffix.lower() == ".xls":
        result["errors"].append(
            "旧版 .xls 格式暂不支持解析（当前使用 openpyxl），请转换为 .xlsx 后重试"
        )
        return result

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:  # noqa: BLE001
        result["errors"].append(f"无法打开 Excel 文件：{e}")
        return result

    try:
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            result["records"].extend(
                _rows_to_records(rows, file_path.name, f"{ws.title}!")
            )
    except Exception as e:  # noqa: BLE001
        result["errors"].append(f"解析 Excel 时出错：{e}")
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass
    return result


# ---------------------------------------------------------------------------
# Word 解析
# ---------------------------------------------------------------------------
def parse_word(file_path):
    file_path = Path(file_path)
    result = _new_result(file_path, "Word")
    if file_path.suffix.lower() == ".doc":
        result["errors"].append(
            "旧版 .doc 格式暂不支持解析（仅支持 .docx），请转换为 .docx 后重试"
        )
        return result

    try:
        doc = Document(str(file_path))
    except Exception as e:  # noqa: BLE001
        result["errors"].append(f"无法打开 Word 文件：{e}")
        return result

    try:
        # 1) 表格
        for t_idx, table in enumerate(doc.tables, start=1):
            rows = [[cell.text for cell in row.cells] for row in table.rows]
            if _looks_like_form_table(rows):
                # 横向表单表：提取「标签|值」键值对
                mapped = {}
                for row in rows:
                    for i in range(0, len(row) - 1, 2):
                        field = match_field(_to_text(row[i]))
                        value = _to_text(row[i + 1])
                        if field and value and field not in mapped:
                            mapped[field] = value
                if mapped:
                    result["records"].append(
                        _build_record(mapped, file_path.name, f"表格{t_idx}表单", {"原文": str(rows)})
                    )
            else:
                result["records"].extend(
                    _rows_to_records(rows, file_path.name, f"表格{t_idx}")
                )

        # 2) 段落
        for p_idx, para in enumerate(doc.paragraphs, start=1):
            text = para.text.strip()
            if not text:
                continue
            fields = _extract_fields_from_text(text)
            if fields:
                result["records"].append(
                    _build_record(fields, file_path.name, f"段落{p_idx}", {"原文": text})
                )
            else:
                result["raw_texts"].append(text)
    except Exception as e:  # noqa: BLE001
        result["errors"].append(f"解析 Word 时出错：{e}")
    return result


# ---------------------------------------------------------------------------
# PDF 解析
# ---------------------------------------------------------------------------
def parse_pdf(file_path):
    file_path = Path(file_path)
    result = _new_result(file_path, "PDF")
    if not _HAS_PYMUPDF:
        result["errors"].append(
            "未安装 PyMuPDF，无法解析 PDF，请执行 pip install -r requirements.txt"
        )
        return result

    try:
        doc = fitz.open(str(file_path))
    except Exception as e:  # noqa: BLE001
        result["errors"].append(f"无法打开 PDF 文件：{e}")
        return result

    try:
        for page_idx, page in enumerate(doc, start=1):
            # 1) 表格
            try:
                table_finder = page.find_tables()
                tables = table_finder.tables
            except Exception as e:  # noqa: BLE001
                tables = []
                result["warnings"].append(f"第{page_idx}页表格识别失败：{e}")
            for t_idx, tab in enumerate(tables, start=1):
                try:
                    rows = tab.extract()
                except Exception as e:  # noqa: BLE001
                    result["warnings"].append(f"第{page_idx}页表格{t_idx}提取失败：{e}")
                    continue
                result["records"].extend(
                    _rows_to_records(rows, file_path.name, f"第{page_idx}页表格{t_idx}")
                )

            # 2) 文字
            try:
                text = page.get_text("text")
            except Exception:  # noqa: BLE001
                text = ""
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                fields = _extract_fields_from_text(line)
                if fields:
                    result["records"].append(
                        _build_record(fields, file_path.name, f"第{page_idx}页文本", {"原文": line})
                    )
                else:
                    result["raw_texts"].append(f"[第{page_idx}页] {line}")
    except Exception as e:  # noqa: BLE001
        result["errors"].append(f"解析 PDF 时出错：{e}")
    finally:
        try:
            doc.close()
        except Exception:  # noqa: BLE001
            pass
    return result


# ---------------------------------------------------------------------------
# 图片解析（本批次仅识别，不做 OCR）
# ---------------------------------------------------------------------------
def parse_image(file_path):
    result = _new_result(file_path, "图片")
    result["warnings"].append("图片文件暂不进行 OCR 识别（后续批次实现）")
    return result


# ---------------------------------------------------------------------------
# 分发与目录聚合
# ---------------------------------------------------------------------------
_HANDLERS = {
    ".xlsx": parse_excel,
    ".xls": parse_excel,
    ".docx": parse_word,
    ".doc": parse_word,
    ".pdf": parse_pdf,
    ".jpg": parse_image,
    ".jpeg": parse_image,
    ".png": parse_image,
}


def parse_file(file_path):
    """按扩展名分发解析单个文件，返回统一结构。"""
    file_path = Path(file_path)
    handler = _HANDLERS.get(file_path.suffix.lower())
    if handler is None:
        result = _new_result(file_path, "未知")
        result["errors"].append(f"不支持的文件类型：{file_path.suffix.lower()}")
        return result
    return handler(file_path)


def parse_directory(directory):
    """递归解析目录下所有支持的文件，返回聚合结果。"""
    directory = Path(directory)
    aggregate = {
        "directory": str(directory),
        "files": [],
        "records": [],
        "raw_texts": [],
        "errors": [],
        "warnings": [],
        "files_parsed": 0,
        "files_failed": 0,
    }
    if not directory.exists():
        aggregate["errors"].append(f"目录不存在：{directory}")
        return aggregate

    for path in sorted(directory.rglob("*")):
        if not path.is_file():
            continue
        handler = _HANDLERS.get(path.suffix.lower())
        if handler is None:
            continue

        rel = path.relative_to(directory)
        file_result = handler(path)
        # 用相对路径覆盖原始文件名，保留目录层级信息
        for rec in file_result["records"]:
            rec["source_file"] = rel.as_posix()

        aggregate["files"].append(file_result)
        aggregate["records"].extend(file_result["records"])
        aggregate["raw_texts"].extend(file_result["raw_texts"])

        if file_result["errors"]:
            aggregate["files_failed"] += 1
            for e in file_result["errors"]:
                aggregate["errors"].append(f"{rel.as_posix()}: {e}")
        else:
            aggregate["files_parsed"] += 1
        for w in file_result["warnings"]:
            aggregate["warnings"].append(f"{rel.as_posix()}: {w}")

    return aggregate
