# 大学生志愿服务智能统计工具

一款基于 **Python + Streamlit** 的桌面 Web 工具，用于智能化统计大学生志愿服务时长：从上传材料、解析识别、名单匹配、人工审核，到按班级导出与模板一致的 Excel 汇总表。

> 当前版本：**v1.0.0**

## 功能总览

| 模块 | 说明 |
|------|------|
| 文件上传与解压 | 批量上传 ZIP，安全解压（防路径穿越/大小限制），识别 Excel/Word/PDF/图片，支持中文文件名 |
| 文档解析 | 读取 Excel（多表）、Word（段落+表格）、PDF（文字+表格，需 PyMuPDF），转为统一记录 |
| OCR 识别 | 图片与扫描 PDF 的中文 OCR（RapidOCR/PaddleOCR 可插拔），保留置信度并标记人工审核 |
| 名单匹配 | 上传班级名单，学号优先精确匹配、学号缺失按姓名匹配（处理空格/全半角），同名不自动认定，支持人工指定 |
| 统计审核 | 按学生统计项目/次数/总时长，小时分钟换算，疑似重复检测，缺失时长补录，排除重复后自动重算 |
| Excel 导出 | 按模板复制样式，动态生成项目列，统计班级总时长/每月人均，多班批量导出并打包 ZIP |

## 技术栈

- Python 3.9+
- Streamlit、Pandas
- OpenPyXL（Excel 读写/导出）
- python-docx（Word 解析）
- PyMuPDF（PDF 解析与渲染）
- PaddleOCR / RapidOCR（OCR，按需安装）

## 目录结构

```
VolunteerStatistics/
├── app.py                 # 主入口（Streamlit 页面）
├── requirements.txt       # 依赖清单
├── start.bat              # Windows 一键启动
├── README.md              # 项目说明
├── TEST_REPORT.md         # 测试报告
├── config/
│   ├── __init__.py
│   └── settings.py        # 全局配置
├── core/
│   ├── file_processor.py  # ZIP 解压与类型识别
│   ├── document_parser.py # 多格式文档解析
│   ├── ocr_processor.py   # OCR 识别
│   ├── student_matcher.py # 名单管理与匹配
│   ├── statistics.py      # 统计与审核
│   └── excel_exporter.py  # Excel 模板导出
├── modules/  utils/       # 预留包
├── uploads/               # 上传文件（程序运行时自动创建）
├── outputs/               # 导出结果
└── temp/                  # 解压临时文件
```

## 环境要求

- Windows 10 / 11
- Python 3.9 及以上（安装时**务必勾选 "Add Python to PATH"**）

## 依赖安装

### 方式一：自动安装（推荐）

直接双击 `start.bat`，脚本会自动创建虚拟环境 `.venv` 并安装 `requirements.txt` 中全部核心依赖。

### 方式二：手动安装核心依赖

```bat
cd /d "项目路径\VolunteerStatistics"
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 安装 OCR（可选，按需）

OCR 依赖较大、不随核心依赖安装，需要识别图片/扫描 PDF 时再装：

```bat
REM 方案 A：RapidOCR（推荐，Windows 兼容性最好，中文模型自动下载）
pip install rapidocr-onnxruntime

REM 方案 B：PaddleOCR（2.x 经典 API）
pip install paddlepaddle
pip install "paddleocr>=2.6,<3.0"
```

## 每次统计志愿服务时的操作步骤

每次有新的志愿服务材料需要统计时，按以下步骤操作：

1. **启动工具**：双击 `start.bat`，浏览器自动打开 http://localhost:8501（未自动打开就手动访问）。
2. **添加班级**：在「🏫 班级名称」区输入本次要统计的班级（可一次输入多个，用逗号/顿号分隔）。
3. **上传名单**：在「🎯 名单管理与自动匹配」区上传各班级学生名单 Excel，并为每个名单选择对应班级（或自动识别）。
4. **上传材料**：在「📁 上传资料」区上传志愿服务材料 ZIP（含 Excel/Word/PDF/图片），系统自动解压并解析。
5. **处理匹配异常**：查看「匹配成功/待人工确认/无法匹配」，对同名、无法匹配的记录人工指定正确学生。
6. **审核补录**：在「📈 统计与审核」区查看学生统计总览；在「缺失时长」补录时长，在「疑似重复」排除重复记录（修改后自动重算）。
7. **导出结果**：在「📄 Excel 导出」区上传真实模板，勾选要导出的班级，点「导出 Excel」，再点「下载 ZIP 打包」。
8. **停止工具**：回到启动窗口按 Ctrl+C（或直接关窗口）。

> 说明：上传的原始文件保存在 `uploads/`、解压到 `temp/`、导出结果在 `outputs/`；原始模板与名单文件**不会被修改**。

## 已知限制

- 旧版二进制格式 `.xls` / `.doc` 暂不支持（请转换为 `.xlsx` / `.docx`）。
- 时长解析暂不支持中文数字（如「二小时」）等口语表达。
- 同名学生（无学号）在统计时可能合并，名单应尽量提供学号。
- OCR 结果需人工审核后再用于统计；Word/PDF 自由文本同一人信息可能拆成多条记录（`原始位置` 可追溯）。
