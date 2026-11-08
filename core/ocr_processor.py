# -*- coding: utf-8 -*-
"""
中文图片与扫描 PDF 的 OCR 识别模块。

保存路径：VolunteerStatistics/core/ocr_processor.py

设计要点：
  - 可插拔 OCR 引擎：优先 RapidOCR（rapidocr-onnxruntime，Windows 兼容性最好），
    其次 PaddleOCR（paddleocr 2.x 经典 API）。也可通过 OCR_BACKEND 强制指定。
  - 支持 JPG / JPEG / PNG 图片。
  - 支持扫描版 PDF（用 PyMuPDF 渲染成图片后再 OCR）。
  - 优先中文识别模型（lang=ch / 中文 onnx 模型）。
  - 提取姓名 / 学号 / 班级 / 项目名称 / 服务日期 / 服务时长等字段。
  - 保留 OCR 原始文本与每字段的置信度。
  - 低置信度字段单独标记；所有 OCR 结果一律标记 review_required=True，必须进入人工审核。
  - 识别失败时记录错误，不崩溃。

安装（二选一，按需安装，不影响项目其他功能）：

  方案 A（推荐，Windows 兼容性最好）：
      pip install rapidocr-onnxruntime

  方案 B（PaddleOCR 2.x 经典 API）：
      pip install paddlepaddle
      pip install "paddleocr>=2.6,<3.0"

  PDF 渲染依赖 PyMuPDF（已在 requirements.txt 中）：
      pip install PyMuPDF
"""
import importlib.util
import os
import re
import tempfile
from collections import OrderedDict
from pathlib import Path

try:
    import fitz  # PyMuPDF
    _HAS_PYMUPDF = True
except Exception:  # noqa: BLE001
    fitz = None
    _HAS_PYMUPDF = False

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
# OCR 引擎选择："auto"（RapidOCR 优先，其次 PaddleOCR）/ "rapidocr" / "paddleocr"
OCR_BACKEND = "auto"

# 置信度阈值：低于该值的字段标记为低置信度，需重点人工审核
CONFIDENCE_THRESHOLD = 0.75

# 扫描 PDF 渲染 DPI（越高越清晰，但越慢）
OCR_DPI = 200


class OcrUnavailableError(Exception):
    """OCR 引擎或依赖不可用。"""


# ---------------------------------------------------------------------------
# 统一字段与别名（与 document_parser 保持一致，供下游复用）
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

FIELD_ALIASES = OrderedDict(
    [
        ("name", ["姓名", "名字", "学生姓名", "志愿者姓名"]),
        ("student_id", ["学号", "学生学号", "学籍号", "工号", "编号"]),
        ("class_name", ["班级", "班级名称", "所属班级", "行政班"]),
        ("project_name", ["项目名称", "服务项目", "活动名称", "志愿项目", "服务内容", "活动", "项目"]),
        ("service_date", ["服务日期", "活动日期", "志愿日期", "服务时间", "活动时间", "日期"]),
        ("service_hours", ["服务时长", "志愿时长", "服务小时", "时长", "小时数", "工时"]),
    ]
)


# ---------------------------------------------------------------------------
# 引擎检测与初始化（懒加载）
# ---------------------------------------------------------------------------
_ENGINE = None  # None=未检测；False=检测过但不可用；对象=可用引擎
_BACKEND_NAME = None
_INIT_ERRORS = []


def _rapidocr_installed():
    return importlib.util.find_spec("rapidocr_onnxruntime") is not None


def _paddleocr_installed():
    return importlib.util.find_spec("paddleocr") is not None


def _init_backend():
    global _ENGINE, _BACKEND_NAME
    if OCR_BACKEND == "rapidocr":
        order = ["rapidocr"]
    elif OCR_BACKEND == "paddleocr":
        order = ["paddleocr"]
    else:
        order = ["rapidocr", "paddleocr"]

    for name in order:
        try:
            if name == "rapidocr":
                from rapidocr_onnxruntime import RapidOCR
                _ENGINE = RapidOCR()
                _BACKEND_NAME = "RapidOCR"
            else:
                from paddleocr import PaddleOCR
                _ENGINE = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
                _BACKEND_NAME = "PaddleOCR"
            return
        except Exception as e:  # noqa: BLE001
            _INIT_ERRORS.append(f"{name}: {e}")
    _ENGINE = False
    _BACKEND_NAME = None


def _get_engine():
    if _ENGINE is None:
        _init_backend()
    return _ENGINE if _ENGINE else None


def get_backend_name():
    """返回当前使用的 OCR 引擎名，未初始化则先检测。"""
    if _ENGINE is None:
        _init_backend()
    return _BACKEND_NAME


def check_backend():
    """检测 OCR 引擎状态，返回摘要信息（不实际加载模型）。"""
    _get_engine()
    return {
        "available": bool(_ENGINE),
        "backend": _BACKEND_NAME,
        "rapidocr_installed": _rapidocr_installed(),
        "paddleocr_installed": _paddleocr_installed(),
        "pymupdf_installed": _HAS_PYMUPDF,
        "init_errors": list(_INIT_ERRORS),
    }


# ---------------------------------------------------------------------------
# 结果归一化（不同引擎 -> 统一行格式）
# ---------------------------------------------------------------------------
def _normalize_rapidocr(result):
    """RapidOCR: ([box, text, score], ...) -> [{text, confidence, box}]"""
    lines = []
    if not result:
        return lines
    for item in result:
        try:
            box, text, score = item[0], item[1], item[2]
            lines.append({"text": str(text).strip(), "confidence": float(score), "box": box})
        except Exception:  # noqa: BLE001
            continue
    return lines


def _normalize_paddleocr(result):
    """PaddleOCR 2.x: [[[box, (text, score)], ...], ...] -> 统一行格式"""
    lines = []
    if not result:
        return lines
    for page in result:
        if not page:
            continue
        for item in page:
            try:
                box, (text, score) = item[0], item[1]
                lines.append({"text": str(text).strip(), "confidence": float(score), "box": box})
            except Exception:  # noqa: BLE001
                continue
    return lines


# ---------------------------------------------------------------------------
# OCR 执行
# ---------------------------------------------------------------------------
def _ocr_image(image_path, engine):
    if _BACKEND_NAME == "RapidOCR":
        result, _elapse = engine(str(image_path))
        return _normalize_rapidocr(result)
    # PaddleOCR
    result = engine.ocr(str(image_path), cls=True)
    return _normalize_paddleocr(result)


def ocr_image(image_path):
    """对单张图片执行 OCR，返回统一行格式列表。"""
    engine = _get_engine()
    if not engine:
        raise OcrUnavailableError(
            "未检测到可用的 OCR 引擎，请先安装 rapidocr-onnxruntime 或 paddleocr+paddlepaddle"
        )
    return _ocr_image(image_path, engine)


def _render_page_to_image(page, out_path, dpi=OCR_DPI):
    pix = page.get_pixmap(dpi=dpi)
    pix.save(out_path)


def ocr_pdf(pdf_path, dpi=OCR_DPI):
    """处理 PDF：文本页直接返回文本，扫描页渲染后 OCR。

    返回列表，每项:
        {"page": 页码, "scanned": bool, "text": str, "ocr_lines": [...]}
    """
    if not _HAS_PYMUPDF:
        raise OcrUnavailableError("未安装 PyMuPDF，无法处理 PDF，请执行 pip install PyMuPDF")
    engine = _get_engine()
    if not engine:
        raise OcrUnavailableError(
            "未检测到可用的 OCR 引擎，请先安装 rapidocr-onnxruntime 或 paddleocr+paddlepaddle"
        )

    doc = fitz.open(str(pdf_path))
    pages = []
    tmp_files = []
    try:
        for page in doc:
            text = (page.get_text("text") or "").strip()
            if len(text) >= 10:
                pages.append(
                    {"page": page.number + 1, "scanned": False, "text": text, "ocr_lines": []}
                )
                continue
            # 扫描页：渲染 + OCR
            fd, tmp_path = tempfile.mkstemp(suffix=".png", prefix="ocr_page_")
            os.close(fd)
            tmp_files.append(tmp_path)
            _render_page_to_image(page, tmp_path, dpi)
            pages.append(
                {
                    "page": page.number + 1,
                    "scanned": True,
                    "text": "",
                    "ocr_lines": _ocr_image(tmp_path, engine),
                }
            )
    finally:
        doc.close()
        for p in tmp_files:
            try:
                os.remove(p)
            except OSError:  # noqa: BLE001
                pass
    return pages


# ---------------------------------------------------------------------------
# 字段提取与记录构建
# ---------------------------------------------------------------------------
def extract_fields_from_lines(lines):
    """从 OCR 行中按「字段名：值」提取字段，返回 (mapped, confidences, raw_text)。"""
    mapped = {}
    confidences = {}
    raw_text = "\n".join(
        str(line.get("text", "")).strip()
        for line in lines
        if str(line.get("text", "")).strip()
    )

    for line in lines:
        text = str(line.get("text", "")).strip()
        conf = line.get("confidence")
        if not text:
            continue
        for field, aliases in FIELD_ALIASES.items():
            if field in mapped:
                continue
            for alias in aliases:
                m = re.search(re.escape(alias) + r"\s*[:：=]\s*([^\s,，;；、|]+)", text)
                if m:
                    mapped[field] = m.group(1).strip(" 　：:=，,；;")
                    confidences[field] = conf
                    break
    return mapped, confidences, raw_text


def _build_ocr_record(mapped, confidences, source_file, source_location, raw_text):
    """构造 OCR 统一记录：保留原始文本与置信度，低置信度标记，强制人工审核。"""
    record = {f: "" for f in RECORD_FIELDS}
    field_confidences = {f: None for f in RECORD_FIELDS}
    low_conf = []

    for f in RECORD_FIELDS:
        v = mapped.get(f)
        record[f] = v if v else ""
        c = confidences.get(f)
        field_confidences[f] = c
        if v and c is not None and float(c) < CONFIDENCE_THRESHOLD:
            low_conf.append(f)

    record["source_file"] = source_file
    record["source_location"] = source_location
    record["raw"] = raw_text
    record["field_confidences"] = field_confidences
    record["low_confidence_fields"] = low_conf
    record["review_required"] = True   # OCR 结果必须进入人工审核
    record["source_type"] = "ocr"
    return record


# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------
def process_file(file_path):
    """对图片或 PDF 执行 OCR 并提取字段，返回统一结构。"""
    file_path = Path(file_path)
    ext = file_path.suffix.lower()
    result = {
        "file": str(file_path),
        "file_type": None,
        "records": [],
        "raw_texts": [],
        "errors": [],
        "warnings": [],
        "backend": None,
    }

    if ext in (".jpg", ".jpeg", ".png"):
        result["file_type"] = "图片"
        try:
            lines = ocr_image(file_path)
            result["backend"] = get_backend_name()
            mapped, confs, raw = extract_fields_from_lines(lines)
            if raw:
                result["raw_texts"].append(raw)
            if mapped:
                result["records"].append(
                    _build_ocr_record(mapped, confs, file_path.name, "图片OCR", raw)
                )
            else:
                result["warnings"].append(
                    "未从图片中识别出可用的「姓名/项目/日期/时长」字段，已保留原始 OCR 文本供人工审核"
                )
        except OcrUnavailableError as e:
            result["errors"].append(str(e))
        except Exception as e:  # noqa: BLE001
            result["errors"].append(f"图片 OCR 失败：{e}")

    elif ext == ".pdf":
        result["file_type"] = "PDF"
        try:
            pages = ocr_pdf(file_path)
            result["backend"] = get_backend_name()
            for p in pages:
                if p["scanned"]:
                    mapped, confs, raw = extract_fields_from_lines(p["ocr_lines"])
                    if raw:
                        result["raw_texts"].append(f"[第{p['page']}页]\n{raw}")
                    if mapped:
                        result["records"].append(
                            _build_ocr_record(
                                mapped, confs, file_path.name, f"第{p['page']}页OCR", raw
                            )
                        )
                    else:
                        result["warnings"].append(
                            f"第{p['page']}页未识别出可用字段，已保留原始 OCR 文本供人工审核"
                        )
                else:
                    result["warnings"].append(
                        f"第{p['page']}页含文本层（非扫描页），请使用文档解析功能处理"
                    )
        except OcrUnavailableError as e:
            result["errors"].append(str(e))
        except Exception as e:  # noqa: BLE001
            result["errors"].append(f"PDF OCR 失败：{e}")

    else:
        result["errors"].append(f"不支持 OCR 的文件类型：{ext}")

    return result
