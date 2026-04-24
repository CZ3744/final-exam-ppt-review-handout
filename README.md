# final-exam-ppt-review-handout

`final-exam-ppt-review-handout` is an OpenClaw/ClawHub-style skill and standalone Python CLI for Chinese university students who need to turn lecture PPTX files into **final-exam-oriented review outlines, Word handouts, and PDFs**.

The project is intentionally designed as an **LLM-orchestrated skill**:

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

## Target scenario

A typical Chinese college final exam workflow:

1. A teacher gives students one or more lecture PPT files.
2. The student asks an LLM to analyze the PPTs instead of mechanically copying slide text.
3. The LLM uses this skill to extract slide content and then writes a structured review outline.
4. The skill renders the final result as Word and PDF files.

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
- A deterministic one-pass fallback mode for environments where no LLM analysis step is available.

## Installation

```bash
pip install -e .
```

PDF export needs LibreOffice or `soffice` on `PATH`. If LibreOffice is unavailable, DOCX generation still works and the report will explain the PDF export failure.

## Recommended LLM-orchestrated workflow

### 1. Extract PPTX into LLM-readable workspace

```bash
python -m ppt_review_handout.cli extract \
  --input ./course_ppts \
  --workspace ./workspace \
  --config examples/sample_config.json
```

Output:

```text
workspace/
├─ extracted/
│  ├─ 第一章 绪论.slides.json
│  ├─ 第一章 绪论.compact.md
│  ├─ 第二章 xxx.slides.json
│  └─ 第二章 xxx.compact.md
├─ report.md
└─ report.json
```

### 2. Let the calling LLM analyze and write handout JSON

The calling LLM should read each `*.compact.md` and, when needed, the matching `*.slides.json`. It should then create one `*.handout.json` per chapter in this schema:

```json
{
  "chapter_title": "第一章 绪论",
  "source_file": "第一章 绪论.pptx",
  "review_goals": ["掌握本章核心概念与考试重点。"],
  "knowledge_framework": ["食品包装科学", "包装基础知识", "包装与社会环境"],
  "core_points": {
    "食品包装科学": ["这里写经过理解和合并后的知识点。"]
  },
  "terms": {
    "食品包装": "采用适当包装材料、容器和包装技术，使食品在运输和贮藏过程中保持价值和原有状态。"
  },
  "comparison_tables": [
    {
      "title": "销售包装与运输包装比较",
      "headers": ["项目", "销售包装", "运输包装"],
      "rows": [["核心作用", "促销与便利消费", "保护与便于物流"]]
    }
  ],
  "processes": {
    "纸的制作流程": ["木料去皮", "切削", "蒸煮", "打浆", "成纸"]
  },
  "exam_points": ["名词解释：食品包装。", "简答：包装的主要功能。"],
  "confusing_points": ["区分销售包装和运输包装。"],
  "quick_summary": ["本章重点是包装的定义、功能、分类和绿色包装理念。"],
  "slide_count": 72,
  "image_heavy_slides": [15, 16]
}
```

### 3. Render handout JSON into Word/PDF

```bash
python -m ppt_review_handout.cli render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --export-pdf \
  --zip-word
```

Output:

```text
outputs/
├─ docx/
│  ├─ 第一章 绪论_复习讲义版.docx
│  └─ ...
├─ pdf/
│  ├─ 第一章 绪论_复习讲义版.pdf
│  └─ ...
├─ word_zip/
│  └─ review_handouts_docx.zip
├─ report.md
└─ report.json
```

## One-pass fallback mode

For quick local tests, this skill also includes a deterministic fallback analyzer:

```bash
python -m ppt_review_handout.cli build \
  --input ./course_ppts \
  --output ./outputs \
  --mode handout \
  --export-pdf \
  --zip-word \
  --keep-intermediate
```

This fallback is useful for smoke tests and rough drafts, but the intended high-quality workflow is `extract -> calling LLM analyzes -> render`.

## OpenClaw / ClawHub usage

The required skill entry file is `SKILL.md`. The skill should be used when the user asks to convert lecture PPTs into final-exam review outlines, review handouts, Word documents, and PDFs.

The LLM using this skill must not treat extraction output as the final answer. It must perform the analysis step itself.

## Development

```bash
pip install -e . pytest
pytest -q
```

## License

MIT-0.
