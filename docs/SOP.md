# SOP: Teacher PPT to Final Exam Review Handout

## Goal

Turn Chinese university course PPT files into final-exam-oriented review handouts, not mechanical slide text dumps.

## Recommended workflow

### 1. Extract

```bash
python -m ppt_review_handout.cli extract --input ./ppts --workspace ./workspace
```

Read:

```text
workspace/extracted/*.compact.md
workspace/extracted/*.slides.json
```

### 2. Analyze as the calling LLM

For each chapter, create a `*.handout.json` file. The content should be organized as:

- review goals;
- knowledge framework;
- core points;
- terms and definitions;
- classification and comparison tables;
- principles, mechanisms, and processes;
- likely exam points;
- confusing-point distinctions;
- quick summary.

Rules:

- Do not mechanically copy PPT text.
- Merge repeated points across pages.
- Preserve definitions, classifications, processes, formulas, and tables.
- Convert tables and diagrams into useful review structures when possible.
- Mark image-heavy pages for review when their meaning cannot be extracted from text.
- Remove headers, footers, teacher/school repetition, template text, and low-value transition pages.

### 3. Render

```bash
python -m ppt_review_handout.cli render --analysis ./workspace/analysis --output ./outputs --export-pdf --zip-word
```

Expected outputs:

```text
outputs/docx/*.docx
outputs/pdf/*.pdf
outputs/word_zip/review_handouts_docx.zip
outputs/report.md
outputs/report.json
```

## Fallback

`build` mode can generate rough rule-based handouts, but it is not the main intended workflow.
