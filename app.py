# -*- coding: utf-8 -*-
"""
大学生志愿服务智能统计工具 · 主入口

保存路径：VolunteerStatistics/app.py

已完成：
  第 1 批：基础框架（标题、工作流程、上传入口、班级录入）
  第 2 批：文件上传与解压模块（ZIP 安全解压、类型识别与统计）
  第 3 批：多格式文档解析模块（Excel/Word/PDF 解析为统一记录）
  第 4 批：OCR 识别模块（独立于 core/ocr_processor.py）
  第 5 批：班级名单管理与自动匹配
  第 6 批：志愿服务统计与人工审核
  第 7 批：Excel 模板生成模块
暂未实现：OCR 审核 UI、Word/PDF 报告导出。
"""
import hashlib
import re
from pathlib import Path

import pandas as pd
import streamlit as st

from config import settings
from core import document_parser
from core import excel_exporter
from core import file_processor
from core import statistics
from core.student_matcher import StudentMatcher, normalize_class, record_signature

# ---------------------------------------------------------------------------
# 1. 页面配置（必须是第一个 Streamlit 命令）
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title=settings.PAGE_TITLE,
    page_icon=settings.PAGE_ICON,
    layout=settings.LAYOUT,
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# 2. 自定义样式
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    html, body, [class*="css"], .stApp {
        font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans SC", sans-serif;
    }
    .block-container { max-width: 1120px; padding-top: 2rem; }
    .title-row { display: flex; align-items: center; gap: 0.6rem; }
    .title-emoji { font-size: 2.2rem; }
    .main-title {
        font-size: 2.1rem; font-weight: 700;
        background: linear-gradient(120deg, #1e3a8a, #3b82f6 55%, #8b5cf6);
        -webkit-background-clip: text; background-clip: text;
        -webkit-text-fill-color: transparent; color: #1e3a8a;
    }
    .subtitle { color: #6b7280; font-size: 1rem; margin: 0.2rem 0 1.4rem; }
    .workflow-container {
        display: flex; gap: 0.8rem; flex-wrap: wrap; margin: 0.8rem 0 1.8rem;
    }
    .step-card {
        flex: 1 1 150px; background: #ffffff; border: 1px solid #e5e7eb;
        border-radius: 14px; padding: 1rem 0.6rem; text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        transition: transform .15s ease, box-shadow .15s ease;
    }
    .step-card:hover { transform: translateY(-3px); box-shadow: 0 10px 22px rgba(59,130,246,0.16); }
    .step-num {
        display: inline-block; background: #eff6ff; color: #2563eb;
        border-radius: 999px; padding: 0.05rem 0.6rem; font-size: 0.72rem;
        font-weight: 600; margin-bottom: 0.35rem;
    }
    .step-icon { font-size: 1.7rem; }
    .step-title { font-weight: 600; color: #1f2937; margin-top: 0.3rem; }
    .step-desc { font-size: 0.78rem; color: #9ca3af; }
    .chip {
        display: inline-block; background: #eff6ff; color: #1d4ed8;
        border: 1px solid #bfdbfe; border-radius: 999px;
        padding: 0.28rem 0.85rem; margin: 0.2rem 0.25rem 0.2rem 0;
        font-size: 0.86rem;
    }
    .footer-note {
        text-align: center; color: #9ca3af; font-size: 0.8rem;
        margin-top: 2rem; padding-top: 1rem; border-top: 1px dashed #e5e7eb;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# 3. 会话状态初始化
# ---------------------------------------------------------------------------
def init_session_state():
    """初始化会话状态变量。"""
    defaults = {
        "class_names": [],
        "saved_files": [],
        "zip_results": {},
        "manual_assignments": {},
        "record_edits": {},
        "excluded_records": {},
        "export_paths": [],
        "export_zip": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_session_state()


# ---------------------------------------------------------------------------
# 4. 工具函数
# ---------------------------------------------------------------------------
def parse_class_names(raw: str):
    """将用户输入按常见分隔符拆分为多个班级名称（去空）。"""
    parts = re.split(r"[,，、;；\s]+", raw.strip())
    return [p for p in parts if p]


def add_class_clicked():
    """「添加班级」按钮回调：写入班级并清空输入框。

    通过 st.button(on_click=...) 触发，在 widget 实例化之前执行，
    因此可以安全修改 st.session_state.new_class_input 而不会触发
    StreamlitWidgetAlreadyInstantiatedError。
    """
    raw = st.session_state.get("new_class_input", "")
    for n in parse_class_names(raw):
        if n not in st.session_state.class_names:
            st.session_state.class_names.append(n)
    st.session_state.new_class_input = ""


def remove_class_clicked():
    """「删除所选」按钮回调：删除选中的班级并清理失效的下拉框状态。"""
    sel = st.session_state.get("remove_class_sel", "")
    if sel and sel in st.session_state.class_names:
        st.session_state.class_names.remove(sel)
    # 选项列表变化后，重置下拉框，避免残留失效值
    st.session_state.pop("remove_class_sel", None)


def clear_classes_clicked():
    """「清空全部班级」按钮回调。"""
    st.session_state.class_names = []
    st.session_state.pop("remove_class_sel", None)
    st.session_state.pop("new_class_input", None)


def format_size(num_bytes: int) -> str:
    """把字节数格式化为可读字符串。"""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / 1024 / 1024:.1f} MB"


def widget_key(key: str) -> str:
    """生成稳定的短小部件键。"""
    return hashlib.md5(key.encode("utf-8")).hexdigest()[:12]


def save_original_file(uploaded_file):
    """把上传的原始文件原样保存到 uploads 目录（不修改），返回保存路径。"""
    settings.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    dest = settings.UPLOAD_DIR / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    if str(dest) not in st.session_state.saved_files:
        st.session_state.saved_files.append(str(dest))
    return dest


def save_roster_file(uploaded_file):
    """把上传的名单 Excel 保存到 uploads/rosters 目录，返回保存路径。"""
    dest_dir = settings.UPLOAD_DIR / "rosters"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    return dest


def get_all_parsed_records():
    """汇总所有已解压 ZIP 中解析出的统一记录。"""
    records = []
    for item in st.session_state.zip_results.values():
        records.extend(item["parse_result"]["records"])
    return records


def student_key(s: dict) -> str:
    """生成学生显示键（用于人工选择下拉框）。"""
    return f"{s['class_name']} · {s['name']}（{s['student_id'] or '无学号'}）"


def build_export_batches(matcher, per_student):
    """按班级分组，返回可导出的批次 [(class_name, students, projects), ...]。"""
    stats_by_class = {}
    for e in per_student:
        cn = e["student"].get("class_name_normalized")
        stats_by_class.setdefault(cn, []).append(e)
    batches = []
    for cls_name in matcher.specified_classes:
        cn = normalize_class(cls_name)
        roster = matcher.get_students(cls_name)
        entries = stats_by_class.get(cn, [])
        if not roster and not entries:
            continue
        students, projects = excel_exporter.build_class_data(roster, entries)
        batches.append((cls_name, students, projects))
    return batches


def handle_zip_uploads(zip_files):
    """处理上传的 ZIP：保存原文件、安全解压、识别类型、解析内容，按指纹去重。"""
    results = st.session_state.get("zip_results", {})
    current_fps = set()

    for f in zip_files:
        fingerprint = hashlib.md5(f"{f.name}:{f.size}".encode("utf-8")).hexdigest()
        current_fps.add(fingerprint)
        if fingerprint in results:
            continue

        orig_path = save_original_file(f)
        dest_dir = settings.EXTRACT_DIR / f"{orig_path.stem}_{fingerprint[:8]}"
        extract_result = file_processor.extract_zip(orig_path, dest_dir)
        parse_result = document_parser.parse_directory(dest_dir)
        results[fingerprint] = {
            "original_name": f.name,
            "size": f.size,
            "result": extract_result,
            "parse_result": parse_result,
        }

    for fp in list(results.keys()):
        if fp not in current_fps:
            del results[fp]

    st.session_state["zip_results"] = results
    return results


def render_zip_result(result: dict):
    """渲染单个压缩包的解压识别结果。"""
    counts = result["counts"]
    st.markdown(f"**识别文件**：共 `{len(result['files'])}` 个")
    st.markdown(
        "".join(
            f'<span class="chip">{t}：{counts[t]}</span>'
            for t in file_processor.TYPE_NAMES
        ),
        unsafe_allow_html=True,
    )
    st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)

    if result["files"]:
        df = pd.DataFrame(result["files"])
        df = df.rename(
            columns={
                "name": "文件名称",
                "type": "文件类型",
                "extension": "扩展名",
                "size": "大小(字节)",
            }
        )
        df["大小"] = df["大小(字节)"].map(format_size)
        df = df[["文件名称", "文件类型", "扩展名", "大小"]]
        st.dataframe(df, hide_index=True, use_container_width=True)
    else:
        st.info("该压缩包中未识别到 Excel / Word / PDF / 图片 文件。")

    if result["errors"]:
        st.warning("⚠️ 以下条目已跳过或处理失败（不影响其余文件）：")
        for e in result["errors"]:
            st.markdown(f"- {e}")


def render_parse_result(parse_result: dict):
    """渲染文档解析得到的统一记录。"""
    records = parse_result.get("records", [])
    st.markdown(f"**解析记录**：共 `{len(records)}` 条")

    if records:
        rows = []
        for r in records:
            rows.append(
                {
                    "姓名": r.get("name", ""),
                    "学号": r.get("student_id", ""),
                    "班级": r.get("class_name", ""),
                    "项目名称": r.get("project_name", ""),
                    "服务日期": r.get("service_date", ""),
                    "服务时长": r.get("service_hours", ""),
                    "原始文件名": r.get("source_file", ""),
                    "原始位置": r.get("source_location", ""),
                }
            )
        df = pd.DataFrame(rows)
        preview = df.head(50)
        st.dataframe(preview, hide_index=True, use_container_width=True)
        if len(df) > 50:
            st.caption(f"仅预览前 50 条，共 {len(df)} 条")
    else:
        st.info("未解析出结构化记录。")

    errors = parse_result.get("errors", [])
    if errors:
        st.warning(f"⚠️ 解析错误（{len(errors)} 条）：")
        for e in errors[:50]:
            st.markdown(f"- {e}")
        if len(errors) > 50:
            st.caption(f"其余 {len(errors) - 50} 条错误已省略")

    warnings = parse_result.get("warnings", [])
    if warnings:
        st.markdown("ℹ️ 提示：")
        for w in warnings[:20]:
            st.markdown(f"- {w}")

    raw_texts = parse_result.get("raw_texts", [])
    if raw_texts:
        with st.expander(f"查看未结构化的原始文本（{len(raw_texts)} 段）"):
            for t in raw_texts[:100]:
                st.markdown(f"- {t}")
            if len(raw_texts) > 100:
                st.caption(f"其余 {len(raw_texts) - 100} 段已省略")


# ---------------------------------------------------------------------------
# 5. 侧边栏
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 📊 志愿服务统计")
    st.markdown(f"**版本**：`{settings.APP_VERSION}`")
    st.divider()
    st.markdown("**当前阶段**")
    st.info("第 7 批 · Excel 导出")
    st.markdown(
        "后续将逐步接入：\n"
        "- OCR 审核 UI\n"
        "- Word/PDF 报告导出"
    )
    st.divider()
    st.markdown("**运行状态**")
    st.markdown(f"- 已上传文件：`{len(st.session_state.saved_files)}` 个")
    recognized = sum(
        len(item["result"]["files"]) for item in st.session_state.zip_results.values()
    )
    st.markdown(f"- 已识别文件：`{recognized}` 个")
    parsed = sum(
        len(item["parse_result"]["records"]) for item in st.session_state.zip_results.values()
    )
    st.markdown(f"- 已解析记录：`{parsed}` 条")
    st.markdown(f"- 已添加班级：`{len(st.session_state.class_names)}` 个")


# ---------------------------------------------------------------------------
# 6. 页头
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="title-row">
        <span class="title-emoji">📊</span>
        <span class="main-title">大学生志愿服务智能统计工具</span>
    </div>
    <div class="subtitle">智能化统计大学生志愿服务时长 · 让材料整理更高效</div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# 7. 工作流程展示
# ---------------------------------------------------------------------------
st.markdown("### 🔄 工作流程")
steps = [
    ("📤", "上传资料", "ZIP / Excel"),
    ("🏫", "配置班级", "多个班级"),
    ("🔍", "智能识别", "OCR 识别材料"),
    ("📈", "统计汇总", "服务时长核算"),
    ("📄", "导出报告", "Excel / Word / PDF"),
]
cards_html = "".join(
    f"""
    <div class="step-card">
        <span class="step-num">第 {i} 步</span>
        <div class="step-icon">{icon}</div>
        <div class="step-title">{title}</div>
        <div class="step-desc">{desc}</div>
    </div>
    """
    for i, (icon, title, desc) in enumerate(steps, start=1)
)
st.markdown(f'<div class="workflow-container">{cards_html}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# 8. 文件上传区域
# ---------------------------------------------------------------------------
st.markdown("### 📁 上传资料")

with st.container(border=True):
    st.markdown("##### 📦 ZIP 文件上传（自动解压 + 解析）")
    st.caption(
        "上传志愿服务材料压缩包（.zip），系统将自动安全解压、识别文件类型，"
        "并解析其中的 Excel / Word / PDF 内容。"
    )
    zip_files = st.file_uploader(
        "选择 ZIP 文件",
        type=["zip"],
        accept_multiple_files=True,
        key="zip_uploader",
        label_visibility="collapsed",
    )
    zip_results = handle_zip_uploads(zip_files or [])

    if zip_files:
        st.success(f"✅ 已接收 {len(zip_files)} 个 ZIP 文件，解压解析完成")
        for f in zip_files:
            fingerprint = hashlib.md5(f"{f.name}:{f.size}".encode("utf-8")).hexdigest()
            item = zip_results.get(fingerprint)
            if not item:
                continue
            with st.expander(
                f"📦 {item['original_name']}（{format_size(item['size'])}）",
                expanded=True,
            ):
                render_zip_result(item["result"])
                st.divider()
                st.markdown("**📄 解析内容（统一记录）**")
                render_parse_result(item["parse_result"])
    else:
        st.info("尚未上传 ZIP 文件。上传后将自动解压、识别类型并解析内容。")

with st.container(border=True):
    st.markdown("##### 📊 Excel 文件上传")
    st.caption("上传班级或学生信息表（.xlsx / .xls）。")
    excel_files = st.file_uploader(
        "选择 Excel 文件",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="excel_uploader",
        label_visibility="collapsed",
    )
    if excel_files:
        for f in excel_files:
            save_original_file(f)
        st.success(f"✅ 已接收 {len(excel_files)} 个 Excel 文件")
        for f in excel_files:
            st.markdown(f"- `{f.name}`（{format_size(f.size)}）")


# ---------------------------------------------------------------------------
# 9. 班级名称录入
# ---------------------------------------------------------------------------
st.markdown("### 🏫 班级名称")
with st.container(border=True):
    st.markdown("**添加班级**")
    st.caption("支持单个添加，也可用逗号（,）、顿号（、）、分号或空格一次输入多个班级。")

    col_input, col_btn = st.columns([4, 1])
    with col_input:
        st.text_input(
            "班级名称",
            key="new_class_input",
            placeholder="例如：计算机2101班、软件2102班、通信2103班",
            label_visibility="collapsed",
        )
    with col_btn:
        st.button(
            "➕ 添加",
            use_container_width=True,
            type="primary",
            on_click=add_class_clicked,
        )

    st.divider()

    if st.session_state.class_names:
        st.markdown("**已添加班级**")
        st.markdown(
            "".join(
                f'<span class="chip">🏫 {n}</span>' for n in st.session_state.class_names
            ),
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)

        col_del_sel, col_del_btn = st.columns([4, 1])
        with col_del_sel:
            st.selectbox(
                "选择要删除的班级",
                options=st.session_state.class_names,
                key="remove_class_sel",
                label_visibility="collapsed",
            )
        with col_del_btn:
            st.button(
                "🗑 删除所选",
                use_container_width=True,
                on_click=remove_class_clicked,
            )

        st.button(
            "🧹 清空全部班级",
            on_click=clear_classes_clicked,
        )
    else:
        st.info("尚未添加班级，请在上方输入班级名称后点击「添加」。")


# ---------------------------------------------------------------------------
# 10. 名单管理与自动匹配
# ---------------------------------------------------------------------------
st.markdown("### 🎯 名单管理与自动匹配")

all_records = get_all_parsed_records()

with st.container(border=True):
    st.markdown("**① 上传班级学生名单（Excel）**")
    st.caption("每个名单可关联到一个班级；也可让系统从名单中的「班级」列自动识别。")
    roster_files = st.file_uploader(
        "选择名单 Excel 文件",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="roster_uploader",
        label_visibility="collapsed",
    )

    roster_assignments = {}
    if roster_files and st.session_state.class_names:
        st.markdown("**② 关联班级**")
        for f in roster_files:
            options = ["（自动识别班级）"] + st.session_state.class_names
            roster_assignments[f.name] = st.selectbox(
                f"「{f.name}」属于哪个班级？",
                options,
                key=f"roster_cls_{f.name}",
            )
            if roster_assignments[f.name] == "（自动识别班级）":
                roster_assignments[f.name] = None
    elif roster_files and not st.session_state.class_names:
        st.warning("请先在上方「班级名称」区域添加要统计的班级，再上传名单。")

    matcher = StudentMatcher(st.session_state.class_names)
    for sig, sdict in st.session_state.manual_assignments.items():
        matcher.manual[sig] = sdict

    load_reports = []
    if roster_files and st.session_state.class_names:
        for f in roster_files:
            path = save_roster_file(f)
            load_reports.append((f.name, matcher.load_roster(path, roster_assignments.get(f.name))))

    for name, rep in load_reports:
        if rep["ok"]:
            msg = f"✅ 「{name}」加载 {rep['students_added']} 名学生"
            if rep["skipped"]:
                msg += f"，跳过 {rep['skipped']} 行"
            st.caption(msg)
        else:
            st.error(f"❌ 「{name}」加载失败：{rep['errors']}")

    results = matcher.match_records(all_records)

    if not st.session_state.class_names:
        st.info("请先在「班级名称」区域添加要统计的班级。")
    elif not all_records:
        st.info("暂无可匹配的记录，请先在「上传资料」区上传并解析 ZIP。")
    elif not matcher.get_students():
        st.warning("已添加班级，但尚未加载任何指定班级的学生名单。")
    else:
        summary = matcher.summarize(results)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("匹配成功", summary["matched"])
        c2.metric("待人工确认", summary["ambiguous"])
        c3.metric("无法匹配", summary["unmatched"])
        c4.metric("已排除", summary["excluded"])

        matched = [r for r in results if r["status"] == "matched"]
        if matched:
            st.markdown("**✅ 匹配成功**")
            rows = []
            match_key_label = {"student_id": "学号", "name": "姓名", "manual": "人工"}
            for r in matched:
                s = r["matched_student"]
                rec = r["record"]
                rows.append(
                    {
                        "姓名": s["name"],
                        "学号": s["student_id"],
                        "班级": s["class_name"],
                        "项目名称": rec.get("project_name", ""),
                        "服务日期": rec.get("service_date", ""),
                        "服务时长": rec.get("service_hours", ""),
                        "匹配方式": match_key_label.get(r["match_key"], r["match_key"]),
                        "来源文件": rec.get("source_file", ""),
                    }
                )
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

        abnormal = [r for r in results if r["status"] in ("unmatched", "ambiguous")]
        if abnormal:
            st.markdown("**⚠️ 异常列表（无法匹配 / 同名学生，需人工处理）**")
            all_students = matcher.get_students()
            for r in abnormal:
                rec = r["record"]
                with st.container(border=True):
                    st.markdown(
                        f"**记录**：姓名「{rec.get('name', '')}」 · "
                        f"学号「{rec.get('student_id', '')}」 · "
                        f"班级「{rec.get('class_name', '')}」"
                    )
                    st.caption(
                        f"来源：{rec.get('source_file', '')} · "
                        f"{rec.get('source_location', '')} · 原因：{r['reason']}"
                    )
                    candidates = r["candidates"] if r["candidates"] else all_students
                    if not candidates:
                        st.warning("名单中没有可选的候选学生。")
                        continue
                    opt_map = {student_key(s): s for s in candidates}
                    options = list(opt_map.keys())
                    sel = st.selectbox(
                        "选择正确学生",
                        options,
                        key=f"manual_sel_{r['index']}",
                    )
                    if st.button("✅ 确认选择", key=f"manual_btn_{r['index']}"):
                        st.session_state.manual_assignments[record_signature(rec)] = opt_map[sel]
                        st.rerun()


# ---------------------------------------------------------------------------
# 11. 统计与审核
# ---------------------------------------------------------------------------
st.markdown("### 📈 统计与审核")

stats_records = statistics.build_stats_records(
    results,
    edits=st.session_state.record_edits,
    excluded=st.session_state.excluded_records,
)
per_student, missing_hours = statistics.compute_statistics(stats_records)
cats = statistics.categorize(results, stats_records)

if not st.session_state.class_names:
    st.info("请先添加班级并加载名单。")
elif not all_records:
    st.info("暂无可统计的记录，请先在「上传资料」区上传并解析 ZIP。")
elif not per_student:
    st.info("暂无匹配成功的记录可统计。")
else:
    st.markdown("**📋 学生统计总览**")
    rows = []
    for e in per_student:
        s = e["student"]
        rows.append(
            {
                "姓名": s["name"],
                "学号": s["student_id"],
                "班级": s["class_name"],
                "服务项目": "、".join(e["projects"]),
                "服务次数": e["count"],
                "总时长": statistics.format_hours(e["total_hours"]),
                "缺失时长(条)": e["hours_missing"],
            }
        )
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    for e in per_student:
        s = e["student"]
        with st.expander(
            f"👤 {s['name']}（{s['student_id'] or '无学号'}）— "
            f"{e['count']} 次 / {statistics.format_hours(e['total_hours'])}"
        ):
            for sr in e["records"]:
                rec = sr["record"]
                dur = (
                    statistics.format_hours(sr["hours"])
                    if sr["hours"] is not None
                    else "⚠️ 缺失"
                )
                st.markdown(
                    f"- 「{rec.get('project_name', '')}」 {rec.get('service_date', '')} · "
                    f"时长「{rec.get('service_hours', '') or '空'}」→ {dur} · "
                    f"来源 `{rec.get('source_file', '')}` / `{rec.get('source_location', '')}`"
                )

    st.divider()
    st.markdown("**🔍 审核中心**")
    tab_missing, tab_dup, tab_unmatch = st.tabs(["缺失时长", "疑似重复", "无法匹配 / 同名"])

    with tab_missing:
        if missing_hours:
            for sr in missing_hours:
                rec = sr["record"]
                rid = statistics.record_id(rec)
                with st.container(border=True):
                    st.markdown(
                        f"**{rec.get('name', '')}**（学号 {sr['student']['student_id']}）· "
                        f"项目「{rec.get('project_name', '')}」"
                    )
                    st.caption(
                        f"来源：{rec.get('source_file', '')} · {rec.get('source_location', '')} · "
                        f"原值：「{rec.get('service_hours', '')}」"
                    )
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        new_hours = st.text_input(
                            "服务时长",
                            key=f"edit_hours_{widget_key(rid)}",
                            placeholder="例如：2小时 或 90分钟",
                        )
                    with c2:
                        if st.button("💾 保存", key=f"save_hours_{widget_key(rid)}"):
                            if new_hours.strip():
                                st.session_state.record_edits[rid] = {
                                    "service_hours": new_hours.strip()
                                }
                            st.rerun()
        else:
            st.success("没有缺失时长的记录。")

    with tab_dup:
        if cats["duplicates"]:
            for g in cats["duplicates"]:
                st.markdown(
                    f"**疑似重复**：学生 `{g['key'][0]}` · 项目 `{g['key'][1]}` · 日期 `{g['key'][2]}`"
                )
                for sr in g["records"]:
                    rec = sr["record"]
                    rid = statistics.record_id(rec)
                    with st.container(border=True):
                        st.markdown(
                            f"时长「{rec.get('service_hours', '') or '空'}」· "
                            f"来源 `{rec.get('source_file', '')}` / `{rec.get('source_location', '')}`"
                        )
                        if st.button("🚫 标记为重复/排除", key=f"excl_{widget_key(rid)}"):
                            st.session_state.excluded_records[rid] = "疑似重复"
                            st.rerun()
        else:
            st.success("没有疑似重复记录。")

        if st.session_state.excluded_records:
            st.markdown("**已排除记录**（可恢复）")
            for rid, reason in list(st.session_state.excluded_records.items()):
                c1, c2 = st.columns([4, 1])
                c1.markdown(f"- `{rid}`（{reason}）")
                if c2.button("↩️ 恢复", key=f"restore_{widget_key(rid)}"):
                    del st.session_state.excluded_records[rid]
                    st.rerun()

    with tab_unmatch:
        st.markdown(
            f"无法匹配：`{len(cats['unmatched'])}` 条 · 同名待确认：`{len(cats['ambiguous'])}` 条"
        )
        st.caption("请在「名单管理与自动匹配」区为这些记录人工指定正确学生。")
        for r in cats["ambiguous"]:
            rec = r["record"]
            st.markdown(
                f"- 同名「{rec.get('name', '')}」学号「{rec.get('student_id', '')}」· "
                f"来源 `{rec.get('source_file', '')}` / `{rec.get('source_location', '')}`"
            )
        for r in cats["unmatched"]:
            rec = r["record"]
            st.markdown(
                f"- 无法匹配「{rec.get('name', '')}」学号「{rec.get('student_id', '')}」· "
                f"原因：{r['reason']}"
            )


# ---------------------------------------------------------------------------
# 12. Excel 导出
# ---------------------------------------------------------------------------
st.markdown("### 📄 Excel 导出")

with st.container(border=True):
    st.markdown("**① 上传 Excel 模板**")
    st.caption("上传你提供的真实模板文件（.xlsx），导出时将完整复制其样式。")
    template_file = st.file_uploader(
        "选择模板文件", type=["xlsx"], key="template_uploader", label_visibility="collapsed"
    )

    template_path = None
    if template_file is not None:
        dest = settings.UPLOAD_DIR / "template" / template_file.name
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(template_file.getbuffer())
        template_path = dest

    batches = build_export_batches(matcher, per_student)

    if template_path is None:
        st.info("请先上传 Excel 模板文件。")
    elif not batches:
        st.info("暂无可导出的班级（请先添加班级、上传名单并完成匹配）。")
    else:
        class_options = [b[0] for b in batches]
        selected = st.multiselect(
            "选择要导出的班级", class_options, default=class_options, key="export_classes_sel"
        )
        if selected and st.button("🚀 导出 Excel", type="primary"):
            sel_batches = [b for b in batches if b[0] in selected]
            out_dir = settings.OUTPUT_DIR
            paths = excel_exporter.export_classes(str(template_path), sel_batches, out_dir)
            zip_path = excel_exporter.zip_files(paths, out_dir / "志愿服务统计导出.zip")
            st.session_state.export_paths = paths
            st.session_state.export_zip = str(zip_path)
            st.rerun()

        if st.session_state.get("export_paths"):
            st.success(f"✅ 已导出 {len(st.session_state.export_paths)} 个文件")
            for p in st.session_state.export_paths:
                st.markdown(f"- `{Path(p).name}`")
            zip_path = st.session_state.get("export_zip")
            if zip_path and Path(zip_path).exists():
                with open(zip_path, "rb") as fh:
                    st.download_button(
                        "⬇️ 下载 ZIP 打包",
                        data=fh.read(),
                        file_name=Path(zip_path).name,
                        mime="application/zip",
                        type="primary",
                    )


# ---------------------------------------------------------------------------
# 13. 页脚说明
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="footer-note">'
    "⚠️ 当前为第 7 批「Excel 导出」版本，OCR 审核 UI、Word/PDF 报告导出将在后续批次逐步上线。"
    "</div>",
    unsafe_allow_html=True,
)
