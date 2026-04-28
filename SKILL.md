---
name: final-exam-ppt-review-handout
description: Extract PPTX/PPTM files into structured intermediate data and render caller-authored handouts into Word/PDF deliverables.
version: 0.3.1
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

# PPT Review Handout Workflow

Use this skill when the user wants one or more PPTX or PPTM decks converted into structured handout deliverables. Legacy binary PPT files are detected and reported, but they must be converted to PPTX first because python-pptx cannot parse them.

This is an LLM-orchestrated skill. The skill does not call a model internally. The calling LLM reads extracted PPT content, understands the material, chooses the organization, merges repeated points, interprets tables and processes, and writes the final handout JSON. The CLI handles deterministic extraction, validation, rendering, packaging, and reports.

## Required workflow

### Step 1: Extract PPTX/PPTM content

```bash
ppt-review-handout extract --input INPUT_PATH --workspace WORKSPACE_DIR --config examples/sample_config.json
```

Module form:

```bash
python -m ppt_review_handout.cli_generic extract --input INPUT_PATH --workspace WORKSPACE_DIR --config examples/sample_config.json
```

This creates extracted compact Markdown files, fuller slides JSON files, and report files inside the workspace.

Read compact.md first. Consult slides.json when table rows, notes, or visual-heavy warnings need more context. If the report warns that many slides depend on visuals, inspect the original PPT with a vision or OCR step before authoring the handout.

### Step 2: Author handout JSON

Before writing handouts, decide how the uploaded materials should be organized. Do not assume every file is a chapter, and do not blindly trust filename order. Use the user request, file names, extracted slide titles, and actual content to infer the best grouping.

For every deck or logical topic, create a file ending in .handout.json. The render step intentionally refuses raw slides.json, so do not skip this semantic handout-writing step.

Rules for the calling LLM:

1. Do not mechanically copy slide text.
2. Do not turn PPT bullets into one-line fragments without synthesis.
3. Preserve definitions, classifications, principles, processes, formulas, table relationships, useful examples, and task-relevant focus points.
4. Compress repeated headers, footers, template text, transition pages, and low-value prompts.
5. Convert real tables into semantic comparison tables.
6. Convert processes into ordered steps.
7. Mark visual-heavy slides in image_heavy_slides when visual review is needed.
8. Do not invent conclusions unsupported by the PPT. Small explanations must stay conservative and source-derived.
9. If file order is ambiguous, choose a sensible order and briefly note the assumption outside the JSON or in the final response/report.

The handout schema is documented in schemas/handout.schema.json. Required fields include chapter_title, source_file, review_goals, knowledge_framework, core_points, terms, comparison_tables, processes, exam_points, confusing_points, quick_summary, slide_count, and image_heavy_slides.

The field names keep backward compatibility with earlier versions; section display names can be changed in config.

### Step 3: Render Word and PDF

```bash
ppt-review-handout render --analysis ANALYSIS_DIR --output OUTPUT_DIR --export-pdf --zip-word
```

Module form:

```bash
python -m ppt_review_handout.cli_generic render --analysis ANALYSIS_DIR --output OUTPUT_DIR --export-pdf --zip-word
```

The render command creates DOCX files, optional PDF files, an optional DOCX zip, and reports.

The default layout is review-margin, which leaves a right-side note area separated by a vertical line. Use --layout standard for a normal full-width document.

## Smoke-test fallback

For smoke tests only:

```bash
ppt-review-handout build --input INPUT_PATH --output OUTPUT_DIR --export-pdf --zip-word --keep-intermediate
```

This uses deterministic rules only. It is not a substitute for semantic LLM analysis. Fallback outputs are marked with generated_by_fallback=true and warnings in the DOCX/report.

## Quality bar

A successful result should read like a synthesized handout, not a slide dump. It should preserve source-supported concepts, tables, processes, distinctions, and summaries while avoiding hallucinated conclusions. The organization should be chosen by the calling LLM based on the material and user request, not hardcoded file names.
