@echo off
setlocal
python -m ppt_review_handout.cli extract --input "%~1" --workspace "%~2"
echo Now ask your LLM to write *.handout.json into %~2\analysis
python -m ppt_review_handout.cli render --analysis "%~2\analysis" --output "%~2\outputs" --export-pdf --zip-word
endlocal
