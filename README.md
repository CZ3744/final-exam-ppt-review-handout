# final-exam-ppt-review-handout

面向中国高校学生期末复习场景的 OpenClaw / ClawHub 风格 skill 与 Python CLI：把老师给的课程 PPTX 解析成结构化中间材料，再由调用方 LLM 理解、重组为考前复习大纲，最后渲染成 Word 讲义和 PDF。

`final-exam-ppt-review-handout` is an OpenClaw/ClawHub-style skill and standalone Python CLI for Chinese university students who need to turn lecture PPTX files into **final-exam-oriented review outlines, Word handouts, and PDFs**.

---

## 中文说明

### 这个项目是做什么的？

很多高校课程期末复习时，老师通常会给一整套 PPT。学生真正需要的不是“逐页复制 PPT 文字”，而是：

- 按章节整理出的复习大纲；
- 每章核心概念、定义、分类、原理和流程；
- 表格、流程图、结构图的语义化整理；
- 易考点、易混淆点和速记总结；
- 可直接打印或分享的 Word / PDF 讲义。

本项目就是为这个场景设计的。它不是泛泛的“PPT 总结器”，而是一个面向 **中国高校期末考试复习** 的 PPT 复习讲义生成 skill。

### 核心设计理念

本项目采用 **调用方 LLM 编排分析** 的设计：

```text
OpenClaw / Claude Code / Codex / ChatGPT 等调用方 LLM
        ↓
使用本 skill 提取 PPTX，生成 slides.json + compact.md
        ↓
调用方 LLM 自己阅读提取结果，理解课程内容并写 handout.json
        ↓
本 skill 把 handout.json 渲染成 DOCX + PDF + ZIP + 报告
```

也就是说：

- 本 skill **不内置 OpenAI / Claude / OpenRouter / Ollama 调用**；
- 真正的内容理解由正在调用它的 LLM 完成；
- 本 skill 负责工程化环节：PPT 提取、中间结构、Word 排版、PDF 导出、打包、质检报告；
- 这样可以避免重复消耗模型额度，也不会绑定某一家模型供应商。

### 它能做什么？

当前能力包括：

- 批量读取单个 PPTX 文件或一个文件夹中的多个 PPTX；
- 按 `第一章`、`第二章`、`chapter 1`、`01` 等文件名自然排序；
- 提取每页：
  - 标题；
  - 正文文本；
  - 表格；
  - 图片/图示数量；
  - 页面类型，如目录页、表格页、图示页、正文页、过渡页；
- 生成两类中间文件：
  - `*.slides.json`：完整结构化数据；
  - `*.compact.md`：适合 LLM 阅读的压缩版文本；
- 接收调用方 LLM 写好的 `*.handout.json`；
- 渲染为：
  - 每章一份 Word `.docx`；
  - 每章一份 PDF；
  - 所有 Word 打包成 zip；
  - `report.md` 和 `report.json` 质检报告；
- 提供 `build` 粗略兜底模式，用于快速烟测。

### 它不能做什么？

为了保持开源、可泛化和供应商无关，本项目默认不做这些事：

- 不在内部调用商业 LLM API；
- 不保证理解纯图片型流程图或截图型表格；
- 不直接处理旧版 `.ppt` 二进制格式，建议先转成 `.pptx`；
- 不把 PPT 原图重新嵌入讲义；
- 不替代调用方 LLM 的语义分析。

如果 PPT 中有大量图片、扫描件、截图表格，调用方 LLM 应结合原 PPT 或视觉/OCR 工具复核。

---

## 推荐工作流

### 第 0 步：安装

```bash
pip install -e .
```

依赖：

```text
python-pptx
python-docx
pypdf
```

如果需要 PDF 导出，请安装 LibreOffice，并确保 `libreoffice` 或 `soffice` 在系统 PATH 中。没有 LibreOffice 时，Word 仍可正常生成，报告会提示 PDF 导出失败原因。

---

### 第 1 步：提取 PPT 内容

把老师给的 PPTX 放进一个目录，例如：

```text
course_ppts/
├─ 第一章 绪论.pptx
├─ 第二章 纸包装材料及容器.pptx
├─ 第三章 塑料包装材料及包装容器.pptx
└─ ...
```

运行：

```bash
python -m ppt_review_handout.cli extract \
  --input ./course_ppts \
  --workspace ./workspace \
  --config examples/sample_config.json
```

输出：

```text
workspace/
├─ extracted/
│  ├─ 第一章 绪论.slides.json
│  ├─ 第一章 绪论.compact.md
│  ├─ 第二章 纸包装材料及容器.slides.json
│  ├─ 第二章 纸包装材料及容器.compact.md
│  └─ ...
├─ report.md
└─ report.json
```

其中：

- `slides.json` 给程序和高级调试使用；
- `compact.md` 给调用方 LLM 阅读和分析使用。

---

### 第 2 步：调用方 LLM 进行真正的复习化分析

这一步是整个项目的关键。

调用方 LLM 应阅读每章的 `compact.md`，必要时查看同名 `slides.json`，然后自己生成 `handout.json`。

重要原则：

- 不要机械复制 PPT 原文；
- 不要把 PPT 每页内容简单换行；
- 要按考试复习逻辑重组；
- 要保留定义、分类、特点、优缺点、原理、流程、公式、表格关系；
- 要压缩课堂互动页、过渡页、页眉页脚、学校名称和模板文字；
- 要主动提炼易考点和易混淆点。

每章建议生成一个：

```text
workspace/analysis/第一章 绪论.handout.json
```

JSON 示例：

```json
{
  "chapter_title": "第一章 绪论",
  "source_file": "第一章 绪论.pptx",
  "review_goals": [
    "掌握食品包装的定义、功能、分类和绿色包装理念。"
  ],
  "knowledge_framework": [
    "食品包装科学",
    "食品包装基础知识",
    "包装与社会和环境",
    "食品包装系统"
  ],
  "core_points": {
    "食品包装的功能": [
      "包装最基本的功能是保护商品，使食品在贮运、销售和消费过程中减少光、氧、水分、温度、微生物和机械冲击等因素造成的品质下降。",
      "包装还具有方便流通与消费、传递信息、促进销售和提高商品价值等作用。"
    ]
  },
  "terms": {
    "食品包装": "采用适当的包装材料、容器和包装技术，把食品包裹起来，使食品在运输和贮藏过程中保持其价值和原有状态。",
    "绿色包装": "能够循环再生利用或降解、节约资源和能源，并在全生命周期中尽量减少对人体健康和环境危害的适度包装。"
  },
  "comparison_tables": [
    {
      "title": "销售包装与运输包装比较",
      "headers": ["比较项目", "销售包装", "运输包装"],
      "rows": [
        ["主要目的", "保护商品、促进销售、方便消费", "保护商品、方便储运和装卸"],
        ["常见形式", "瓶、罐、盒、袋等小包装", "瓦楞纸箱、木箱、托盘、集装箱等"]
      ]
    }
  ],
  "processes": {
    "PPT 复习整理逻辑": [
      "读取 PPT 原始内容",
      "识别章节结构和主要知识点",
      "合并重复内容",
      "提炼易考点和易混淆点",
      "输出 Word 和 PDF 复习讲义"
    ]
  },
  "exam_points": [
    "名词解释：食品包装、绿色包装。",
    "简答：食品包装的主要功能。",
    "比较：销售包装和运输包装的区别。"
  ],
  "confusing_points": [
    "包装既可以指容器、材料及辅助物，也可以指实施包装操作的技术活动。",
    "销售包装更侧重消费者和促销，运输包装更侧重物流保护。"
  ],
  "quick_summary": [
    "本章重点是食品包装的定义、功能、分类以及包装与环境之间的关系。",
    "复习时应重点掌握包装的保护功能、促进销售功能和绿色包装 4R1D 原则。"
  ],
  "slide_count": 72,
  "image_heavy_slides": [15, 16]
}
```

---

### 第 3 步：渲染成 Word 和 PDF

运行：

```bash
python -m ppt_review_handout.cli render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --export-pdf \
  --zip-word
```

输出：

```text
outputs/
├─ docx/
│  ├─ 第一章 绪论_复习讲义版.docx
│  ├─ 第二章 纸包装材料及容器_复习讲义版.docx
│  └─ ...
├─ pdf/
│  ├─ 第一章 绪论_复习讲义版.pdf
│  ├─ 第二章 纸包装材料及容器_复习讲义版.pdf
│  └─ ...
├─ word_zip/
│  └─ review_handouts_docx.zip
├─ report.md
└─ report.json
```

---

## 一键兜底模式

如果只是想快速测试工具链是否能跑通，可以使用 `build`：

```bash
python -m ppt_review_handout.cli build \
  --input ./course_ppts \
  --output ./outputs \
  --mode handout \
  --export-pdf \
  --zip-word \
  --keep-intermediate
```

注意：`build` 使用规则生成粗略讲义，只适合烟测和草稿，不等于高质量复习讲义。真正高质量结果应使用：

```text
extract → 调用方 LLM 分析 → render
```

---

## OpenClaw / ClawHub 使用方式

本项目的 skill 入口是：

```text
SKILL.md
```

调用方 LLM 看到用户提出“把课程 PPT 整理成期末复习大纲 / 复习讲义 / Word / PDF”时，应按照 `SKILL.md` 指示执行三段式流程。

关键要求：

```text
不能把 extract 结果直接当最终答案。
必须由调用方 LLM 自己完成复习化分析。
```

---

## 后续示例产物

后续可以在仓库中增加一个 `examples/outputs/` 或 Release 附件，用真实课程 PPT 跑出的 Word/PDF 作为示例。

建议不要把大量二进制 Word/PDF 长期直接提交到 main 分支。更推荐：

- 小体积示例可放 `examples/outputs/`；
- 大体积完整产物放 GitHub Release；
- 或者只提交 `report.md`、`sample.compact.md`、`sample.handout.json`，让用户自己运行生成 Word/PDF。

---

## English overview

This project is intentionally designed as an **LLM-orchestrated skill**:

```text
calling LLM / OpenClaw / Claude Code / Codex
        ↓
uses this skill to extract PPTX into structured slides.json + compact.md
        ↓
calling LLM reads the extracted content and writes exam-review handout.json
        ↓
this skill renders handout.json into DOCX + PDF + QA reports
```

The skill itself does **not** call OpenAI, Claude, OpenRouter, or Ollama. The model that invokes the skill is responsible for understanding the course content, merging knowledge points, interpreting tables, and writing the review handout. This keeps the project vendor-neutral, privacy-friendly, and suitable for open-source reuse.

## Features

- Batch PPTX input: one file or a folder.
- Natural chapter sorting, including Chinese numerals such as `第一章`, `第二章`.
- PPTX extraction to:
  - `*.slides.json` for complete structured data;
  - `*.compact.md` for LLM-friendly reading.
- Caller-authored `*.handout.json` rendering to:
  - Word `.docx` review handouts;
  - PDF files through LibreOffice / `soffice` when available;
  - zipped Word deliverables;
  - `report.md` and `report.json` QA reports.
- A deterministic one-pass fallback mode for smoke tests.

## Installation

```bash
pip install -e .
```

PDF export needs LibreOffice or `soffice` on `PATH`. If LibreOffice is unavailable, DOCX generation still works and the report will explain the PDF export failure.

## Development

```bash
pip install -e . pytest
pytest -q
```

## License

MIT-0.
