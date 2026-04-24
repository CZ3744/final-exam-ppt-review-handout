# Architecture

`final-exam-ppt-review-handout` is an LLM-orchestrated skill. It separates deterministic file processing from semantic course understanding.

## Core idea

```text
PPTX files
  -> extractor
  -> slides.json + compact.md
  -> calling LLM performs semantic analysis
  -> handout.json
  -> renderer
  -> Word + PDF + reports
```

The skill itself does not call OpenAI, Claude, OpenRouter, Gemini, or Ollama. This keeps the repository vendor-neutral and lets the invoking model determine the quality of analysis.

## Layers

1. PPT discovery and chapter sorting
   - Finds PPT/PPTX files.
   - Sorts Arabic and Chinese chapter names.

2. Extraction
   - Reads PPTX files with `python-pptx`.
   - Extracts slide text, tables, image counts, and slide roles.
   - Removes common boilerplate and repeated headers/footers.

3. Compact context generation
   - Converts extracted slides into LLM-readable `compact.md`.
   - Keeps enough information for analysis without forcing the LLM to read huge JSON first.

4. Calling LLM analysis step
   - Reads `compact.md` and optionally `slides.json`.
   - Produces `handout.json` following the schema in `SKILL.md`.
   - This is where true PPT understanding happens.

5. Rendering
   - Renders handout JSON into a structured Chinese Word document.
   - Uses LibreOffice / `soffice` for PDF conversion when available.
   - Writes `report.md` and `report.json` with generated files, warnings, and errors.

6. Fallback mode
   - `build` provides a deterministic rough draft for smoke tests.
   - It is useful for local verification, not the primary high-quality workflow.
