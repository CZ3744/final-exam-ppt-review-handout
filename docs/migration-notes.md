# Migration notes: generic workflow CLI

This project now treats the final-exam Chinese handout style as an optional profile, not as core behavior.

## Active entrypoint

The installed console scripts point to:

```text
ppt_review_handout.cli_generic:main
```

`cli_generic` wraps `workflow_cli` and patches a version-tolerant visual element detector for different `python-pptx` enum sets.

## Compatibility note

Older files such as `cli.py` and `cli_v2.py` may still exist in the source tree for compatibility with older external references. New development should target:

```text
src/ppt_review_handout/workflow_cli.py
src/ppt_review_handout/cli_generic.py
```

## Default vs profile behavior

Default config:

```text
examples/sample_config.json
```

is generic and avoids course-specific or exam-specific wording.

Optional Chinese final-exam profile:

```text
examples/zh_final_exam_config.json
```

restores Chinese section titles, fonts, note-column label, and file suffixes for the original use case.

## Recommended commands

Generic extraction:

```bash
ppt-review-handout extract --input ./ppts --workspace ./workspace --config examples/sample_config.json
```

Chinese final-exam rendering profile:

```bash
ppt-review-handout render --analysis ./workspace/analysis --output ./outputs --config examples/zh_final_exam_config.json --zip-word --export-pdf
```

## Remaining cleanup

After downstream users stop importing old module paths, remove or reduce `cli.py` and `cli_v2.py` to simple compatibility wrappers that re-export from `cli_generic`.
