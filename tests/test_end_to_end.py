import json
from pathlib import Path

from docx import Document
from pptx import Presentation

from ppt_review_handout.cli_v2 import main


def create_demo_pptx(path: Path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "第一章 测试章节"
    box = slide.shapes.add_textbox(914400, 1371600, 6400800, 914400)
    box.text_frame.text = "测试概念是指用于端到端测试的概念。"
    table_shape = slide.shapes.add_table(2, 2, 914400, 2743200, 5486400, 1371600)
    table = table_shape.table
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "说明"
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "B"
    prs.save(path)


def test_extract_then_render_pipeline(tmp_path: Path):
    ppt_dir = tmp_path / "ppts"
    workspace = tmp_path / "workspace"
    outputs = tmp_path / "outputs"
    analysis = workspace / "analysis"
    ppt_dir.mkdir()
    analysis.mkdir(parents=True)
    create_demo_pptx(ppt_dir / "第一章 测试.pptx")

    assert main(["extract", "--input", str(ppt_dir), "--workspace", str(workspace)]) == 0
    assert list((workspace / "extracted").glob("*.slides.json"))
    assert list((workspace / "extracted").glob("*.compact.md"))

    handout = {
        "chapter_title": "第一章 测试章节",
        "source_file": "第一章 测试.pptx",
        "review_goals": ["掌握测试概念。"],
        "knowledge_framework": ["测试概念", "测试表格"],
        "core_points": {"测试概念": ["测试概念用于验证提取和渲染流程。"]},
        "terms": {"测试概念": "用于端到端测试的概念。"},
        "comparison_tables": [
            {"title": "测试对比表", "headers": ["项目", "说明"], "rows": [["A", "B"]]}
        ],
        "processes": {"测试流程": ["创建 PPTX", "提取内容", "渲染 DOCX"]},
        "exam_points": ["简答：说明测试流程。"],
        "confusing_points": ["区分 slides.json 和 handout.json。"],
        "quick_summary": ["端到端流程应稳定生成 DOCX。"],
        "slide_count": 1,
        "image_heavy_slides": [],
    }
    (analysis / "demo.handout.json").write_text(json.dumps(handout, ensure_ascii=False), encoding="utf-8")

    assert main(["render", "--analysis", str(analysis), "--output", str(outputs), "--zip-word"]) == 0
    docx_files = list((outputs / "docx").glob("*.docx"))
    assert len(docx_files) == 1
    assert (outputs / "word_zip" / "review_handouts_docx.zip").exists()
    doc = Document(str(docx_files[0]))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "第一章 测试章节" in text
    assert "测试概念" in text
