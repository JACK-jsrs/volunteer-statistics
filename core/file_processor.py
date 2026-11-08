# -*- coding: utf-8 -*-
"""
文件处理核心模块：ZIP 安全解压与文件类型识别。

保存路径：VolunteerStatistics/core/file_processor.py

功能：
  - 安全解压 ZIP（防路径穿越、限制条目/文件/总量大小，防 zip 炸弹）
  - 支持中文文件名（自动识别 GBK / UTF-8 编码）
  - 识别 Excel / Word / PDF / 图片（JPG/JPEG/PNG）文件
  - 统计各类型文件数量
  - 单个文件损坏只记录错误，不影响其他文件处理
  - 不修改原始压缩包（只读取，不解压覆盖原文件）
"""
import os
import shutil
import stat
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# 常量配置
# ---------------------------------------------------------------------------
# 扩展名 -> 类型分类
SUPPORTED_EXTENSIONS = {
    ".xlsx": "Excel",
    ".xls": "Excel",
    ".docx": "Word",
    ".doc": "Word",
    ".pdf": "PDF",
    ".jpg": "图片",
    ".jpeg": "图片",
    ".png": "图片",
}

# 用于统计展示的类型顺序
TYPE_NAMES = ("Excel", "Word", "PDF", "图片")

# 安全限制
MAX_ENTRIES = 1000                        # 单个压缩包最多条目数
MAX_SINGLE_FILE_SIZE = 100 * 1024 * 1024  # 单个解压文件最大 100 MB
MAX_TOTAL_SIZE = 500 * 1024 * 1024        # 解压总量最大 500 MB
MAX_ZIP_FILE_SIZE = 200 * 1024 * 1024     # 压缩包本身最大 200 MB


class FileProcessError(Exception):
    """文件处理过程中的自定义异常。"""


def _decode_name(name: str) -> str:
    """修复中文文件名乱码（兼容 GBK 与 UTF-8 编码）。"""
    if not name:
        return ""
    # zipfile 在未声明 UTF-8 时默认按 cp437 解码；
    # 若名称包含 cp437 之外的字符，说明本身已是正确的 Unicode，直接返回。
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    for encoding in ("gbk", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return name


def classify_file(filename: str):
    """根据扩展名返回文件类型分类；不支持的类型返回 None。"""
    ext = Path(filename).suffix.lower()
    return SUPPORTED_EXTENSIONS.get(ext)


def _safe_extract_path(dest_dir: Path, member_name: str) -> Path:
    """计算安全的解压目标路径，防止路径穿越（zip slip）。"""
    parts = []
    for part in member_name.replace("\\", "/").split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            raise FileProcessError(f"检测到不安全的路径（包含 ..）：{member_name}")
        if ":" in part:
            raise FileProcessError(f"检测到不安全的路径（包含盘符）：{member_name}")
        parts.append(part)
    if not parts:
        raise FileProcessError(f"无效的条目名称：{member_name}")

    target = dest_dir.joinpath(*parts)
    # 二次校验：解析后的真实路径必须位于解压目录之内
    resolved = Path(os.path.realpath(target))
    base = Path(os.path.realpath(dest_dir))
    if resolved != base and base not in resolved.parents:
        raise FileProcessError(f"路径越界，已拒绝：{member_name}")
    return target


def extract_zip(zip_path, dest_dir):
    """安全解压一个 ZIP，并返回文件识别与统计结果。

    参数:
        zip_path: 压缩包路径（str 或 Path）
        dest_dir: 解压目标目录（str 或 Path）

    返回 dict:
        {
            "zip_name": str,          # 压缩包文件名
            "extracted_dir": str,     # 解压目录
            "files": [ {"name", "type", "extension", "size"}, ... ],
            "counts": {"Excel": n, "Word": n, "PDF": n, "图片": n},
            "extension_counts": {".xlsx": n, ...},
            "errors": [str, ...],
        }
    """
    result = {
        "zip_name": Path(zip_path).name,
        "extracted_dir": str(dest_dir),
        "files": [],
        "counts": {name: 0 for name in TYPE_NAMES},
        "extension_counts": {},
        "errors": [],
    }

    zip_path = Path(zip_path)
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)

    if not zip_path.exists():
        result["errors"].append(f"压缩包不存在：{zip_path.name}")
        return result

    zip_size = zip_path.stat().st_size
    if zip_size > MAX_ZIP_FILE_SIZE:
        result["errors"].append(
            f"压缩包过大（{zip_size / 1024 / 1024:.1f} MB），"
            f"超过限制 {MAX_ZIP_FILE_SIZE / 1024 / 1024:.0f} MB"
        )
        return result

    try:
        zf = zipfile.ZipFile(zip_path)
    except zipfile.BadZipFile as e:
        result["errors"].append(f"压缩包损坏或不是有效的 ZIP 文件：{e}")
        return result
    except Exception as e:  # noqa: BLE001
        result["errors"].append(f"无法打开压缩包：{e}")
        return result

    total_size = 0
    with zf:
        infos = zf.infolist()
        if len(infos) > MAX_ENTRIES:
            result["errors"].append(
                f"压缩包条目过多（{len(infos)} 个），超过限制 {MAX_ENTRIES} 个"
            )
            return result

        for info in infos:
            if info.is_dir():
                continue

            # 跳过符号链接，避免安全风险
            mode = info.external_attr >> 16
            if mode and stat.S_ISLNK(mode):
                result["errors"].append(
                    f"跳过符号链接条目：{_decode_name(info.filename)}"
                )
                continue

            name = _decode_name(info.filename)

            # 单个文件解压大小限制
            if info.file_size > MAX_SINGLE_FILE_SIZE:
                result["errors"].append(
                    f"跳过过大文件：{name}（{info.file_size / 1024 / 1024:.1f} MB，"
                    f"超过 {MAX_SINGLE_FILE_SIZE / 1024 / 1024:.0f} MB）"
                )
                continue

            # 解压总量限制（防 zip 炸弹）
            total_size += info.file_size
            if total_size > MAX_TOTAL_SIZE:
                result["errors"].append(
                    f"解压总量超过限制 {MAX_TOTAL_SIZE / 1024 / 1024:.0f} MB，"
                    f"已停止继续解压"
                )
                break

            # 安全路径校验
            try:
                target = _safe_extract_path(dest_dir, name)
            except FileProcessError as e:
                result["errors"].append(str(e))
                continue

            # 解压单个文件；损坏只记录错误，不影响其他文件
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(info) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst, length=1024 * 1024)
            except (zipfile.BadZipFile, OSError, RuntimeError, ValueError) as e:
                result["errors"].append(f"解压文件失败：{name}（{e}）")
                continue

            # 识别并统计
            ftype = classify_file(name)
            if ftype is not None:
                ext = Path(name).suffix.lower()
                result["files"].append(
                    {
                        "name": name,
                        "type": ftype,
                        "extension": ext,
                        "size": target.stat().st_size,
                    }
                )
                result["counts"][ftype] += 1
                result["extension_counts"][ext] = (
                    result["extension_counts"].get(ext, 0) + 1
                )

    return result
