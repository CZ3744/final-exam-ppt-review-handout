---
name: final-exam-ppt-review-handout
description: Extract Chinese university lecture PPTX/PPTM files and render caller-authored final-exam review handouts into Word/PDF deliverables.
version: 0.2.1
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

Use this skill when the user wants to convert teacher-provided university course `.pptx` / `.pptm` files into final-exam-oriented review outlines, Word handouts, and PDFs. Legacy binary `.ppt` files are detected and reported, but they must be converted to `.pptx` first because `python-pptx` cannot parse them.

This is an **LLM-orchestrated skill**. The skill does not call a model internally. You, the calling LLM, are responsible for reading the extracted PPT content, understanding the course logic, deciding the correct organization order, merging repeated points, interpreting table relationships, and writing the final review handout structure. The skill handles extraction, intermediate files, DOCX/PDF rendering, packaging, and QA reports.

## When to use

Use this skill for requests such as:

- “把老师给的 PPT 整理成期末复习大纲。”
- “把这些课程 PPTX/PPTM 做成考前复习讲义 Word 和 PDF。”
- “不要机械提取文字，要按知识点和易考点整理。”
- “每章一份 Word，每章一份 PDF。”

Do not use it for generic business slide summarization unless the user specifically wants exam-review handouts.

## Required workflow

### Step 1: Extract PPTX/PPTM content

Run either the installed command:

```bash
ppt-review-handout extract \
  --input <pptx-pptm-file-or-directory> \
  --workspace <workspace-dir> \
  --config examples/sample_config.json
```

or the module form:

```bash
python -m ppt_review_handout.cli_v2 extract \
  --input <pptx-pptm-file-or-directory> \
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

The `compact.md` files are optimized for you to read. The `slides.json` files preserve fuller slide structure, including text, tables, detected roles, image counts, and notes. If the report warns that many pages are image-heavy, inspect the original PPT with a vision/OCR-capable step before writing the final handout.

### Step 2: Decide organization and write handout JSON

Before writing handouts, decide how the uploaded materials should be organized. Do not assume every file must be named `第一章`, `第二章`, etc. Use the user's request, file names, extracted slide titles, and actual content to infer the best order and grouping.

Common cases:

1. Continuous course chapters:
   - Example files: `绪论.pptx`, `第一章 材料性能.pptx`, `第二章 晶体结构.pptx`
   - Recommended output: one handout per chapter, ordered by course logic: 绪论 → 第一章 → 第二章.

2. Single final-review deck:
   - Example file: `机械工程材料总复习.pptx`
   - Recommended output: one complete review handout organized by topics inside the deck.

3. Independent topic decks:
   - Example files: `金属材料复习.pptx`, `热处理专题.pptx`, `有色金属专题.pptx`
   - Recommended output: one handout per independent deck. Do not force them into a chapter sequence.

4. Mixed or unclear files:
   - Example files: `绪论.pptx`, `实验复习.pptx`, `期末重点.pptx`, `补充资料.pptx`
   - Recommended behavior: infer a sensible order, document the decision in `report.md` or a short note, and avoid pretending the files form a strict chapter sequence if they do not.

For every chapter or independent deck, read the corresponding `*.compact.md` and, if needed, `*.slides.json`. Then create a `*.handout.json` file. The render step intentionally refuses raw `slides.json`, so do not skip this semantic handout-writing step.

You must follow these rules:

1. Do **not** mechanically copy slide text.
2. Do **not** turn PPT bullets into one-line fragments.
3. Preserve definitions, classifications, principles, processes, formulas, table relationships, examples with exam value, and likely short-answer/essay points.
4. Compress repeated headers, footers, school names, teacher names, template text, transition pages, and low-value classroom prompts.
5. Convert real tables into semantic comparison tables.
6. Convert processes into ordered steps.
7. Mark image-heavy slides in `image_heavy_slides` when visual review is needed.
8. Do not invent professional conclusions not supported by the PPT. If you add a small explanation for clarity, keep it conservative and derived from the slide content.
9. If file order is ambiguous, choose the order that best matches teaching/review logic and briefly note the assumption.

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
ppt-review-handout render \
  --analysis <directory-containing-handout-json> \
  --output <output-dir> \
  --export-pdf \
  --zip-word
```

or:

```bash
python -m ppt_review_handout.cli_v2 render \
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

The default layout is `review-margin`, which leaves a right-side annotation area separated by a vertical line. This is recommended for printed exam-review notes. Use `--layout standard` when the user wants a normal full-width document or a conservative fallback layout.

## Fallback one-pass mode

For smoke tests only, you may run:

```bash
ppt-review-handout build \
  --input <pptx-pptm-file-or-directory> \
  --output <output-dir> \
  --mode handout \
  --export-pdf \
  --zip-word \
  --keep-intermediate
```

This uses deterministic rules and is not a substitute for your semantic analysis. Fallback outputs are explicitly marked as rough drafts in the DOCX/report.

## Quality bar

A successful result should look like a real exam-review handout, not a slide dump. The final deliverables should include Word and PDF versions, clear chapter/topic structure, key concepts, comparison tables, mechanisms/processes, likely exam points, confusing-point distinctions, and a quick final summary. The organization order should be chosen by the calling LLM based on course logic, not blindly forced by file names.
