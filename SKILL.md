---
name: 销售数据智能分析报告生成 (SalesReview)
description: 读取包含销售阶段、金额、工作记录的 Excel 文件，通过 Gemini 大模型多维度分析生成重点项目跟进、团队对比、员工绩效评估等自动化 PDF 报告。
---

# 技能背景 (Background)

该技能封装了基于 Gemini 2.5 Flash 模型的自动化商业智能分析工具，用于从 Fxiaoke 导出的 Excel 销售数据源中，自动化生成精美的 PDF 格式数据分析复盘报告。涵盖三大核心功能：
1. **单人销售报告 (`analyze_person.py`)**：按员工拆分线索进行单人销售行为分析与建议。
2. **片区总览报告 (`analyze_region.py`)**：按大区拆分生成多维度的区域销售评估、行业对标及打分系统。
3. **全国大盘报告 (`analyze_sales.py`)**：全局统筹分析整个公司的商机漏斗。

# 技能使用说明 (Instructions)

## 1. 环境准备与配置
- 确保系统安装了 Python 3.10+。
- 执行依赖安装：`pip install -r requirements.txt`。
- 如果是首次运行，需要安装 Playwright 浏览器依赖：`playwright install chromium`。
- 配置密钥：在根目录创建 `.env` 文件，写入：
  ```env
  GEMINI_API_KEY="您的_GEMINI_KEY"
  ```
- 配置组织架构：复制 `config.example.json` 为 `config.json`，并根据实际人员填写映射表。

## 2. 工具调用方式
系统会自动寻找 `downloads/` 文件夹中最新下载的 Excel 销售数据（也可以通过命令行参数直接指定文件路径）。

### 生成单人复盘报告
```bash
python analyze_person.py [path/to/excel.xlsx]
```

### 生成大区复盘报告
```bash
python analyze_region.py [path/to/excel.xlsx]
```

### 生成全国全局复盘报告
```bash
python analyze_sales.py [path/to/excel.xlsx]
```

## 3. 核心依赖
- **Playwright**：用于高质量的 PDF 渲染排版。
- **google-genai**：调用最新的 Gemini Flash 分析内容。
- **pandas / openpyxl**：高效的 Excel 表格与结构化数据处理。
