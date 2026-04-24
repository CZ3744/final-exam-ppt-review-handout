---
name: final-exam-ppt-review-handout
description: Extract Chinese university lecture PPTX files and render caller-authored final-exam review handouts into Word/PDF deliverables.
version: 0.2.0
metadata:
  openclaw:
    requires:
      bins:
        - python
      anyBins:
        - libreoffice
        - soffice
    emoji: "📚"
    homepage: https://github.com/CZ3744/final-exam-ppt-review-handout
    skillKey: final-exam-ppt-review-handout
---

# Final Exam PPT Review Handout

Use this skill when the user wants to convert teacher-provided university course PPT/PPTX files into final-exam-oriented review outlines, Word handouts, and PDFs.

This is an **LLM-orchestrated skill**. The skill does not call a model internally. You, the calling LLM, are responsible for reading the extracted PPT content, understanding the course logic, merging repeated points, interpreting table relationships, and writing the final review handout structure. The skill handles extraction, intermediate files, DOCX/PDF rendering, packaging, and QA reports.

## When to use

Use this skill for requests such as:

- “把老师给的 PPT 整理成期末复习大纲。”
- “把这些课程 PPT 做成考前复习讲义 Word 和 PDF。”
- “不要机械提取文字，要按知识点和易考点整理。”
- “每章一份 Word，每章一份 PDF。”

Do not use it for generic business slide summarization unless the user specifically wants exam-review handouts.

## Required workflow

### Step 1: Extract PPT content

Run:

```bash
python -m ppt_review_handout.cli extract \
  --input <pptx-file-or-directory> \
  --workspace <workspace-dir> \
  --config examples/sample_config.json
```

This creates:

```text
<workspace-dir>/extracted/*.slides.json
<workspace-dir>/extracted/*.compact.md
<workspace-dir>/report.md
<workspace-dir>/report.json
```

The `compact.md` files are optimized for you to read. The `slides.json` files preserve fuller slide structure, including text, tables, detected roles, image counts, and notes.

### Step 2: You analyze and write handout JSON

For every chapter, read the corresponding `*.compact.md` and, if needed, `*.slides.json`. Then create a `*.handout.json` file.

You must follow these rules:

1. Do **not** mechanically copy slide text.
2. Do **not** turn PPT bullets into one-line fragments.
3. Preserve definitions, classifications, principles, processes, formulas, table relationships, examples with exam value, and likely short-answer/essay points.
4. Compress repeated headers, footers, school names, teacher names, template text, transition pages, and low-value classroom prompts.
5. Convert real tables into semantic comparison tables.
6. Convert processes into ordered steps.
7. Mark image-heavy slides in `image_heavy_slides` when visual review is needed.
8. Do not invent professional conclusions not supported by the PPT. If you add a small explanation for clarity, keep it conservative and derived from the slide content.

Handout JSON schema:

```json
{
  "chapter_title": "第一章 绪论",
  "source_file": "第一章 绪论.pptx",
  "review_goals": [],
  "knowledge_framework": [],
  "core_points": {
    "一级知识点标题": ["整理后的复习知识点"]
  },
  "terms": {
    "术语": "定义解释"
  },
  "comparison_tables": [
    {
      "title": "表格标题",
      "headers": ["项目", "类别A", "类别B"],
      "rows": [["比较项", "内容A", "内容B"]]
    }
  ],
  "processes": {
    "流程名称": ["步骤1", "步骤2"]
  },
  "exam_points": [],
  "confusing_points": [],
  "quick_summary": [],
  "slide_count": 0,
  "image_heavy_slides": []
}
```

### Step 3: Render Word and PDF

Run:

```bash
python -m ppt_review_handout.cli render \
  --analysis <directory-containing-handout-json> \
  --output <output-dir> \
  --export-pdf \
  --zip-word
```

This creates:

```text
<output-dir>/docx/*.docx
<output-dir>/pdf/*.pdf
<output-dir>/word_zip/review_handouts_docx.zip
<output-dir>/report.md
<output-dir>/report.json
```

## Fallback one-pass mode

For smoke tests only, you may run:

```bash
python -m ppt_review_handout.cli build \
  --input <pptx-file-or-directory> \
  --output <output-dir> \
  --mode handout \
  --export-pdf \
  --zip-word \
  --keep-intermediate
```

This uses deterministic rules and is not a substitute for your semantic analysis.

## Quality bar

A successful result should look like a real exam-review handout, not a slide dump. The final deliverables should include Word and PDF versions, clear chapter structure, key concepts, comparison tables, mechanisms/processes, likely exam points, confusing-point distinctions, and a quick final summary.
