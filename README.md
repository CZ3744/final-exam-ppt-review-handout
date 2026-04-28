# final-exam-ppt-review-handout

A generic OpenClaw / LLM workflow skill for turning `.pptx` / `.pptm` slide decks into caller-authored structured handouts.

The project is intentionally split into two responsibilities:

- **The CLI handles deterministic work:** PPTX/PPTM discovery, text/table/notes/visual metadata extraction, intermediate files, handout JSON validation, DOCX/PDF rendering, zipping, and reports.
- **The calling LLM handles semantic work:** reading the extracted content, understanding the material, deciding the organization, merging repeated points, interpreting tables/processes, and authoring `*.handout.json`.

It is not a raw PPT-to-PDF converter and it is not a fully autonomous content generator. The high-quality path is always:

```text
extract -> calling LLM reads compact.md/slides.json -> calling LLM writes *.handout.json -> render
```

Legacy binary `.ppt` files are detected and reported, but must be converted to `.pptx` first because `python-pptx` cannot parse them.

---

## What it does

- Processes one PPTX/PPTM file or a directory of decks.
- Supports recursive discovery with collision-safe output names.
- Extracts slide titles, text blocks, tables, notes, and visual-heavy slide hints.
- Writes compact Markdown summaries plus fuller `slides.json` files.
- Requires caller-authored `*.handout.json` for rendering; raw `slides.json` is refused.
- Renders DOCX handouts with configurable title suffix, filename suffix, section titles, fonts, zip name, and note-column label.
- Optionally exports PDF via LibreOffice/soffice.
- Writes `report.md` and `report.json` with warnings/errors.

The previous China-university/final-exam wording is now only an example profile idea. Core code and default config are generic.

---

## Install

```bash
pip install -e .
```

Development:

```bash
pip install -e . pytest
pytest -q
```

---

## Step 1: Extract

```bash
ppt-review-handout extract \
  --input ./course_ppts \
  --workspace ./workspace \
  --config examples/sample_config.json
```

Recursive:

```bash
ppt-review-handout extract \
  --input ./course_ppts \
  --workspace ./workspace \
  --recursive
```

Outputs:

```text
workspace/extracted/*.compact.md
workspace/extracted/*.slides.json
workspace/report.md
workspace/report.json
```

`compact.md` is optimized for the calling LLM. `slides.json` preserves more structure.

---

## Step 2: Author handout JSON

The calling LLM must read the extracted files and create `workspace/analysis/*.handout.json`.

Required shape:

```json
{
  "chapter_title": "Deck or topic title",
  "source_file": "source.pptx",
  "review_goals": [],
  "knowledge_framework": [],
  "core_points": {
    "section title": ["caller-authored point"]
  },
  "terms": {
    "term": "definition"
  },
  "comparison_tables": [
    {
      "title": "table title",
      "headers": ["item", "A", "B"],
      "rows": [["comparison item", "A content", "B content"]]
    }
  ],
  "processes": {
    "process name": ["step 1", "step 2"]
  },
  "exam_points": [],
  "confusing_points": [],
  "quick_summary": [],
  "slide_count": 0,
  "image_heavy_slides": []
}
```

The schema is documented in `schemas/handout.schema.json`.

Important constraints:

- Do not mechanically copy every slide bullet.
- Do not invent unsupported conclusions.
- Convert real tables into semantic comparison tables.
- Convert processes into ordered steps.
- Mark visual-heavy slides for original-PPT or OCR/vision review.
- Do not assume every file is a chapter; use the user request and extracted content to decide grouping.

---

## Step 3: Render

```bash
ppt-review-handout render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --export-pdf \
  --zip-word
```

Standard full-width layout:

```bash
ppt-review-handout render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --layout standard \
  --zip-word
```

Outputs:

```text
outputs/docx/*.docx
outputs/pdf/*.pdf
outputs/word_zip/handouts_docx.zip
outputs/report.md
outputs/report.json
```

PDF export is best-effort. If LibreOffice/soffice is not installed, DOCX still renders and the report explains why PDF was skipped.

---

## Smoke-test fallback

`build` is intentionally only a deterministic smoke-test path. It proves extraction and rendering work, but it is not a substitute for caller LLM analysis.

```bash
ppt-review-handout build \
  --input ./course_ppts \
  --output ./outputs \
  --keep-intermediate \
  --zip-word
```

Fallback output is marked with `generated_by_fallback=true` and warnings in the report/DOCX.

---

## Configuration

See `examples/sample_config.json` for generic defaults. You can configure:

- `document_title_suffix`
- `output_filename_suffix`
- `zip_filename`
- `note_column_label`
- `body_font`
- `heading_font`
- `remove_patterns`
- `max_table_rows_in_summary`
- `max_text_items_per_slide`
- `text_clip_limit`
- `image_heavy_threshold`
- `absolute_paths`
- `sections`

`remove_patterns` supports exact text and `re:` regular-expression patterns.

---

## Notes

The installed console scripts now point to `ppt_review_handout.workflow_cli:main`. Older `cli.py` / `cli_v2.py` files may remain as compatibility references, but new development should target `workflow_cli.py`.

## License

MIT-0.
