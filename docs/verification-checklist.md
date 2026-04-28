# Verification checklist

Run locally after pulling the latest repository.

Install dependencies:

- pip install -e . pytest

Check commands:

- ppt-review-handout --help
- ppt-review-handout extract --help
- ppt-review-handout render --help
- ppt-review-handout build --help
- python -m ppt_review_handout.cli_generic --help

Run tests:

- pytest -q

Manual smoke test:

1. Place PPTX files in a temporary input folder.
2. Run extract with examples/sample_config.json.
3. Confirm compact Markdown, slides JSON, report Markdown, and report JSON are created.
4. Create an analysis folder with one valid file ending in .handout.json.
5. Run render with examples/sample_config.json.
6. Confirm DOCX output and handouts_docx.zip are created.
7. Run render again with examples/zh_final_exam_config.json.
8. Confirm Chinese profile output suffixes and zip name are used.

Regression expectations:

- Default sample config is generic.
- Chinese final exam wording lives only in the optional profile config.
- Raw slides JSON should not be rendered as a handout.
- Build mode is only a smoke-test fallback.
- Recursive extraction should avoid overwriting files with duplicate stems.
