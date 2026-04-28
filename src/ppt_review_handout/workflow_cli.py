from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

CN_DIGITS = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
CN_UNITS = {"十": 10, "百": 100, "千": 1000}
SUPPORTED_SUFFIXES = {".pptx", ".pptm"}
UNSUPPORTED_SUFFIXES = {".ppt"}
LAYOUT_CHOICES = ("review-margin", "standard")
BOILERPLATE = [
    r"^PowerPoint Template$",
    r"^单击此处编辑母版文本样式$",
    r"^第二级$",
    r"^第三级$",
    r"^第四级$",
    r"^第五级$",
]
REQUIRED_KEYS = {
    "chapter_title", "source_file", "review_goals", "knowledge_framework",
    "core_points", "terms", "comparison_tables", "processes",
    "exam_points", "confusing_points", "quick_summary", "slide_count", "image_heavy_slides",
}
SECTION_TITLES = {
    "review_goals": "Review goals",
    "knowledge_framework": "Knowledge framework",
    "core_points": "Core points",
    "terms": "Terms and definitions",
    "comparison_tables": "Classification and comparison tables",
    "processes": "Principles, mechanisms and processes",
    "exam_points": "Review focus",
    "confusing_points": "Distinctions and possible confusions",
    "quick_summary": "Quick summary",
    "image_heavy_slides": "Slides needing visual review",
}


@dataclass
class SkillConfig:
    language: str = "generic"
    document_title_suffix: str = "Structured handout"
    output_filename_suffix: str = "handout"
    zip_filename: str = "handouts_docx.zip"
    note_column_label: str = "Notes"
    body_font: str = "Arial"
    heading_font: str = "Arial"
    remove_patterns: list[str] = field(default_factory=list)
    max_table_rows_in_summary: int = 8
    max_text_items_per_slide: int = 12
    text_clip_limit: int = 900
    image_heavy_threshold: float = 0.30
    absolute_paths: bool = False
    sections: dict[str, str] = field(default_factory=lambda: dict(SECTION_TITLES))


def load_config(path: str | None, warnings: list[str] | None = None) -> SkillConfig:
    cfg = SkillConfig()
    if not path:
        return cfg
    p = Path(path).expanduser()
    if not p.exists():
        msg = f"Config file not found: {p}"
        if warnings is not None:
            warnings.append(msg)
            return cfg
        raise FileNotFoundError(msg)
    data = json.loads(p.read_text(encoding="utf-8"))
    for key, value in data.items():
        if hasattr(cfg, key):
            setattr(cfg, key, value)
    cfg.sections = {**SECTION_TITLES, **(data.get("sections") or {})}
    return cfg


def chinese_to_int(text: str) -> int:
    text = text.strip()
    if text.isdigit():
        return int(text)
    total = 0
    current = 0
    unit_seen = False
    for ch in text:
        if ch in CN_DIGITS:
            current = CN_DIGITS[ch]
        elif ch in CN_UNITS:
            unit_seen = True
            total += (current or 1) * CN_UNITS[ch]
            current = 0
    total += current
    if not unit_seen and total == 0:
        for ch in text:
            total = total * 10 + CN_DIGITS.get(ch, 0)
    return total


def chapter_index(name: str) -> int:
    m = re.search(r"第\s*([零〇一二两三四五六七八九十百千0-9]+)\s*[章节]", name)
    if m:
        return chinese_to_int(m.group(1))
    m = re.search(r"chapter\s*([0-9]+)", name, re.I)
    if m:
        return int(m.group(1))
    m = re.search(r"(^|[^0-9])([0-9]{1,3})([^0-9]|$)", name)
    return int(m.group(2)) if m else 9999


def safe_name(name: str) -> str:
    for ch in '<>:"/\\|?*':
        name = name.replace(ch, "_")
    return name.strip() or "untitled"


def unique_stem(path: Path, root: Path | None = None) -> str:
    base = safe_name(path.stem)
    rel = str(path if root is None else path.relative_to(root) if path.is_relative_to(root) else path)
    digest = hashlib.sha1(rel.encode("utf-8", errors="ignore")).hexdigest()[:8]
    return f"{base}_{digest}"


def relpath(path: str | Path, base: Path, absolute: bool = False) -> str:
    p = Path(path)
    if absolute:
        return str(p)
    try:
        return str(p.relative_to(base))
    except Exception:
        return str(p.name)


def discover_pptx(path: Path, recursive: bool = False) -> tuple[list[Path], list[Path]]:
    if not path.exists():
        return [], []
    if path.is_file():
        suffix = path.suffix.lower()
        if suffix in SUPPORTED_SUFFIXES:
            return [path], []
        if suffix in UNSUPPORTED_SUFFIXES:
            return [], [path]
        return [], []
    iterator = path.rglob("*") if recursive else path.iterdir()
    candidates = [p for p in iterator if p.is_file() and not p.name.startswith("~$")]
    supported = [p for p in candidates if p.suffix.lower() in SUPPORTED_SUFFIXES]
    unsupported = [p for p in candidates if p.suffix.lower() in UNSUPPORTED_SUFFIXES]
    return sorted(supported, key=lambda p: (chapter_index(p.name), str(p))), sorted(unsupported, key=str)


def normalized(text: str) -> str:
    return " ".join(str(text).split()).strip()


def is_noise(text: str, custom: list[str] | None = None) -> bool:
    t = normalized(text)
    if not t:
        return True
    for pattern in BOILERPLATE + list(custom or []):
        p = str(pattern).strip()
        if not p:
            continue
        if p.startswith("re:"):
            if re.search(p[3:], t):
                return True
        elif re.match(p, t) or normalized(p) == t:
            return True
    return False


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def extract_shape_text(shape) -> list[str]:
    out: list[str] = []
    if getattr(shape, "has_text_frame", False) and shape.text_frame:
        for para in shape.text_frame.paragraphs:
            text = normalized(" ".join(run.text for run in para.runs)) or normalized(getattr(para, "text", ""))
            if text:
                out.append(text)
    return out


def extract_table(shape) -> list[list[str]] | None:
    if not getattr(shape, "has_table", False):
        return None
    return [[normalized(cell.text) for cell in row.cells] for row in shape.table.rows]


def extract_notes(slide) -> str:
    try:
        notes = slide.notes_slide.notes_text_frame
        return "\n".join(p.text.strip() for p in notes.paragraphs if p.text.strip())
    except Exception:
        return ""


def visual_weight(shape) -> int:
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return 1
    if getattr(shape, "has_chart", False):
        return 1
    if getattr(shape, "has_table", False):
        return 1
    if shape.shape_type in {MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.OLE_OBJECT, MSO_SHAPE_TYPE.PLACEHOLDER}:
        return 1
    return 0


def detect_role(title: str, texts: list[str], tables: list[list[list[str]]], visual_count: int) -> str:
    joined = " ".join([title] + texts[:6])
    if tables:
        return "table"
    if visual_count >= 3 and len(texts) <= 5:
        return "visual-heavy"
    if any(k in joined for k in ["目录", "CONTENTS", "Agenda", "Outline"]):
        return "toc-like"
    if len(texts) <= 1 and visual_count == 0:
        return "transition-like"
    return "content"


def extract_presentation(pptx_path: Path, remove_patterns: list[str] | None = None) -> dict:
    prs = Presentation(str(pptx_path))
    slides = []
    removed: list[str] = []
    for idx, slide in enumerate(prs.slides, start=1):
        texts: list[str] = []
        tables: list[list[list[str]]] = []
        visual_count = 0
        for shape in iter_shapes(slide.shapes):
            visual_count += visual_weight(shape)
            table = extract_table(shape)
            if table:
                tables.append(table)
            for text in extract_shape_text(shape):
                if is_noise(text, remove_patterns):
                    removed.append(text)
                else:
                    texts.append(text)
        title = texts[0] if texts and len(texts[0]) <= 80 else f"Slide {idx}"
        slides.append({
            "index": idx,
            "title": title,
            "texts": texts[1:] if texts and texts[0] == title else texts,
            "tables": [{"rows": t} for t in tables],
            "visual_element_count": visual_count,
            "image_count": visual_count,
            "notes": extract_notes(slide),
            "detected_role": detect_role(title, texts, tables, visual_count),
        })
    return {
        "source_file": str(pptx_path),
        "chapter_title": pptx_path.stem,
        "slide_count": len(slides),
        "slides": slides,
        "removed_boilerplate": sorted(set(removed))[:200],
    }


def clip(text: str, limit: int = 900) -> str:
    text = normalized(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def deck_to_compact_md(deck: dict, cfg: SkillConfig) -> str:
    lines = [
        f"# {deck['chapter_title']}", "",
        f"- Source file: {Path(deck['source_file']).name}",
        f"- Slide count: {deck['slide_count']}", "",
        "## Calling LLM workflow constraints", "",
        "Read this compact file and consult the matching slides.json when needed. Author the final handout yourself: merge repeated points, interpret real tables, convert processes into ordered steps, mark visual-heavy slides for review, and avoid mechanically copying slide bullets.", "",
        "Render only caller-authored *.handout.json files. Do not pass raw *.slides.json to the render step.", "",
        "## Slide digest", "",
    ]
    for slide in deck["slides"]:
        lines.append(f"## Slide {slide['index']}: {slide['title']}")
        lines.append(f"- Detected role: {slide['detected_role']}")
        if slide.get("visual_element_count"):
            lines.append(f"- Visual elements: {slide['visual_element_count']} (review original PPT or a vision/OCR step if the slide relies on visuals)")
        if slide.get("texts"):
            lines.append("- Text:")
            for text in slide["texts"][: int(cfg.max_text_items_per_slide)]:
                lines.append(f"  - {clip(text, int(cfg.text_clip_limit))}")
            if len(slide["texts"]) > cfg.max_text_items_per_slide:
                lines.append(f"  - ... {len(slide['texts']) - cfg.max_text_items_per_slide} more items in slides.json")
        if slide.get("tables"):
            lines.append("- Tables:")
            for i, table in enumerate(slide["tables"], start=1):
                rows = table.get("rows", [])
                lines.append(f"  - Table {i}: {len(rows)} rows")
                for row in rows[: int(cfg.max_table_rows_in_summary)]:
                    lines.append("    - " + " | ".join(clip(cell, 80) for cell in row))
        if slide.get("notes"):
            lines.append(f"- Notes: {clip(slide['notes'], int(cfg.text_clip_limit))}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def validate_handout_schema(handout: dict) -> list[str]:
    errors: list[str] = []
    missing = sorted(REQUIRED_KEYS - set(handout))
    if missing:
        errors.append("Missing handout fields: " + ", ".join(missing))
    if "slides" in handout:
        errors.append("This looks like extracted slides.json, not caller-authored handout.json.")
    list_fields = ["review_goals", "knowledge_framework", "exam_points", "confusing_points", "quick_summary", "image_heavy_slides"]
    for field in list_fields:
        if field in handout and not isinstance(handout[field], list):
            errors.append(f"{field} must be a list.")
    for field in ["core_points", "terms", "processes"]:
        if field in handout and not isinstance(handout[field], dict):
            errors.append(f"{field} must be an object.")
    if "comparison_tables" in handout:
        if not isinstance(handout["comparison_tables"], list):
            errors.append("comparison_tables must be a list.")
        else:
            for i, table in enumerate(handout["comparison_tables"], start=1):
                if not isinstance(table, dict) or not isinstance(table.get("headers", []), list) or not isinstance(table.get("rows", []), list):
                    errors.append(f"comparison_tables[{i}] must contain headers:list and rows:list.")
    if "slide_count" in handout and not isinstance(handout["slide_count"], int):
        errors.append("slide_count must be an integer.")
    return errors


def ensure_child(parent, tag: str):
    child = parent.find(qn(tag))
    if child is None:
        child = OxmlElement(tag)
        parent.append(child)
    return child


def set_east_asia_font(element, font_name: str):
    rpr = ensure_child(element, "w:rPr")
    rfonts = ensure_child(rpr, "w:rFonts")
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:ascii"), font_name)


def set_font(run, font="Arial", size=10.5, bold=None):
    run.font.name = font
    set_east_asia_font(run._element, font)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def setup_doc(layout: str, cfg: SkillConfig) -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    if layout == "review-margin":
        sec.top_margin = sec.bottom_margin = Cm(1.5)
        sec.left_margin = sec.right_margin = Cm(1.4)
    else:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1.8)
    styles = doc.styles
    styles["Normal"].font.size = Pt(10.5)
    set_east_asia_font(styles["Normal"]._element, cfg.body_font)
    for name, size in [("Title", 18), ("Heading 1", 15), ("Heading 2", 13), ("Heading 3", 11)]:
        styles[name].font.size = Pt(size)
        styles[name].font.bold = True
        set_east_asia_font(styles[name]._element, cfg.heading_font)
    return doc


def set_cell_border(cell, **edges):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge, attrs in edges.items():
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        for key, value in attrs.items():
            element.set(qn("w:" + key), str(value))


def set_cell_text(cell, text: str, cfg: SkillConfig, bold: bool = False):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(str(text))
    set_font(r, cfg.body_font, 9.5, bold)


def prepare_content_target(doc: Document, layout: str, cfg: SkillConfig):
    if layout == "standard":
        return doc
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    left, right = table.cell(0, 0), table.cell(0, 1)
    left.width, right.width = Cm(12.7), Cm(4.2)
    nil = {"val": "nil"}
    line = {"val": "single", "sz": "8", "space": "0", "color": "BFBFBF"}
    for cell in (left, right):
        set_cell_border(cell, top=nil, start=nil, bottom=nil, end=nil)
    set_cell_border(left, end=line)
    if cfg.note_column_label:
        set_font(right.paragraphs[0].add_run(cfg.note_column_label), cfg.body_font, 8)
    return left


def add_heading(target, text: str, cfg: SkillConfig, level: int = 1):
    if hasattr(target, "add_heading"):
        return target.add_heading(text, level=level)
    p = target.add_paragraph(style=f"Heading {level}")
    set_font(p.add_run(text), cfg.heading_font, 12 if level >= 2 else 14, True)
    return p


def add_bullets(target, items, cfg: SkillConfig, style="List Bullet"):
    for item in items or []:
        if item:
            p = target.add_paragraph(style=style)
            set_font(p.add_run(str(item)), cfg.body_font)


def add_title(target, title: str, cfg: SkillConfig):
    p = target.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run(title), cfg.heading_font, 18, True)
    if cfg.document_title_suffix:
        p2 = target.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(p2.add_run(cfg.document_title_suffix), cfg.heading_font, 12, True)


def add_table(target, title: str, headers: list[str], rows: list[list[str]], layout: str, cfg: SkillConfig):
    if not rows:
        return
    add_heading(target, title or "Comparison table", cfg, level=3)
    width = max(len(headers), max((len(r) for r in rows), default=0), 1)
    if layout == "review-margin" and width >= 4:
        for row in rows:
            add_bullets(target, [row[0] if row else "Item"], cfg)
            for col_idx, cell in enumerate(row[1:], start=1):
                label = headers[col_idx] if col_idx < len(headers) else f"Field {col_idx + 1}"
                p = target.add_paragraph()
                p.paragraph_format.left_indent = Cm(0.6)
                set_font(p.add_run(f"{label}: "), cfg.heading_font, 10.5, True)
                set_font(p.add_run(str(cell)), cfg.body_font)
        return
    table = target.add_table(rows=1, cols=width)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i in range(width):
        set_cell_text(table.rows[0].cells[i], headers[i] if i < len(headers) else f"Field {i + 1}", cfg, True)
    for row in rows:
        cells = table.add_row().cells
        for i in range(width):
            set_cell_text(cells[i], row[i] if i < len(row) else "", cfg)


def handout_to_docx(handout: dict, path: Path, layout: str = "review-margin", cfg: SkillConfig | None = None):
    cfg = cfg or SkillConfig()
    errors = validate_handout_schema(handout)
    if errors:
        raise ValueError("; ".join(errors))
    doc = setup_doc(layout, cfg)
    target = prepare_content_target(doc, layout, cfg)
    add_title(target, handout.get("chapter_title", "Untitled"), cfg)
    meta = target.add_paragraph()
    set_font(meta.add_run(f"Source: {Path(handout.get('source_file', '')).name}; slides: {handout.get('slide_count', 0)}; layout: {layout}"), cfg.body_font, 9)
    if handout.get("generated_by_fallback"):
        p = target.add_paragraph()
        set_font(p.add_run("Warning: generated by deterministic smoke-test fallback; caller LLM review is required for quality output."), cfg.heading_font, 10.5, True)
    add_heading(target, cfg.sections["review_goals"], cfg, 1); add_bullets(target, handout.get("review_goals", []), cfg)
    add_heading(target, cfg.sections["knowledge_framework"], cfg, 1); add_bullets(target, handout.get("knowledge_framework", []), cfg)
    add_heading(target, cfg.sections["core_points"], cfg, 1)
    for title, points in (handout.get("core_points") or {}).items():
        add_heading(target, title, cfg, 2); add_bullets(target, points, cfg)
    add_heading(target, cfg.sections["terms"], cfg, 1)
    for term, definition in (handout.get("terms") or {}).items():
        p = target.add_paragraph(); set_font(p.add_run(f"{term}: "), cfg.heading_font, 10.5, True); set_font(p.add_run(str(definition)), cfg.body_font)
    add_heading(target, cfg.sections["comparison_tables"], cfg, 1)
    for t in handout.get("comparison_tables", []) or []:
        add_table(target, t.get("title", "Table"), t.get("headers", []), t.get("rows", []), layout, cfg)
    add_heading(target, cfg.sections["processes"], cfg, 1)
    for name, steps in (handout.get("processes") or {}).items():
        add_heading(target, name, cfg, 2); add_bullets(target, steps, cfg, "List Number")
    for field in ["exam_points", "confusing_points", "quick_summary"]:
        add_heading(target, cfg.sections[field], cfg, 1); add_bullets(target, handout.get(field, []), cfg)
    if handout.get("image_heavy_slides"):
        add_heading(target, cfg.sections["image_heavy_slides"], cfg, 1)
        add_bullets(target, [f"Slide {i}: review original visual content." for i in handout["image_heavy_slides"]], cfg)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def export_pdf(docx_path: Path, pdf_dir: Path) -> tuple[str | None, str | None]:
    exe = shutil.which("libreoffice") or shutil.which("soffice")
    if not exe:
        return None, "PDF export skipped: libreoffice/soffice not found."
    pdf_dir.mkdir(parents=True, exist_ok=True)
    try:
        proc = subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(docx_path)], check=True, timeout=180, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except subprocess.TimeoutExpired:
        return None, "PDF export failed: LibreOffice conversion timed out."
    except subprocess.CalledProcessError as exc:
        return None, f"PDF export failed: {(exc.stderr or b'').decode('utf-8', errors='ignore')[:300]}"
    pdf = pdf_dir / (docx_path.stem + ".pdf")
    return (str(pdf), None) if pdf.exists() else (None, f"PDF export failed: converted file not found. stdout={(proc.stdout or b'').decode('utf-8', errors='ignore')[:200]}")


def write_report(root: Path, records: list[dict], cfg: SkillConfig | None = None) -> None:
    cfg = cfg or SkillConfig()
    root.mkdir(parents=True, exist_ok=True)
    errors = [e for r in records for e in r.get("errors", [])]
    warnings = [w for r in records for w in r.get("warnings", [])]
    summary = {
        "file_count": len(records),
        "docx_count": sum(1 for r in records if r.get("docx")),
        "pdf_count": sum(1 for r in records if r.get("pdf")),
        "warnings": warnings,
        "errors": errors,
        "records": records,
    }
    (root / "report.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = ["# PPT handout workflow report", "", f"- Files: {summary['file_count']}", f"- DOCX: {summary['docx_count']}", f"- PDF: {summary['pdf_count']}", f"- Warnings: {len(warnings)}", f"- Errors: {len(errors)}", ""]
    for r in records:
        lines += [f"## {r.get('chapter_title')}", ""]
        for key in ["source_file", "compact_md", "intermediate_slides", "docx", "pdf"]:
            if r.get(key): lines.append(f"- {key}: {r[key]}")
        for w in r.get("warnings", []): lines.append(f"- warning: {w}")
        for e in r.get("errors", []): lines.append(f"- error: {e}")
        lines.append("")
    (root / "report.md").write_text("\n".join(lines), encoding="utf-8")


def image_heavy_warning(deck: dict, cfg: SkillConfig) -> str | None:
    count = deck.get("slide_count", 0) or 0
    if not count:
        return None
    heavy = [s for s in deck.get("slides", []) if s.get("detected_role") == "visual-heavy"]
    if len(heavy) / count >= float(cfg.image_heavy_threshold):
        return f"Visual-heavy deck: {len(heavy)}/{count} slides need original-PPT or vision/OCR review."
    return None


def extract_cmd(args) -> int:
    warnings: list[str] = []
    cfg = load_config(args.config, warnings)
    workspace = Path(args.workspace).expanduser().resolve(); workspace.mkdir(parents=True, exist_ok=True)
    input_path = Path(args.input).expanduser().resolve()
    records: list[dict] = []
    if not input_path.exists():
        records.append({"source_file": str(input_path), "chapter_title": "Input not found", "warnings": warnings, "errors": [f"Input path does not exist: {input_path}"]})
        write_report(workspace, records, cfg); return 1
    files, unsupported = discover_pptx(input_path, recursive=args.recursive)
    for old in unsupported:
        records.append({"source_file": relpath(old, workspace, cfg.absolute_paths), "chapter_title": old.stem, "slide_count": 0, "warnings": ["Legacy .ppt is not supported by python-pptx. Convert it to .pptx first."], "errors": []})
    if not files:
        records.append({"source_file": str(input_path), "chapter_title": "No supported files", "warnings": warnings, "errors": ["No supported .pptx/.pptm files found."]})
        write_report(workspace, records, cfg); return 1
    root = input_path if input_path.is_dir() else input_path.parent
    for ppt in files:
        rec = {"source_file": relpath(ppt, workspace, cfg.absolute_paths), "chapter_title": ppt.stem, "slide_count": 0, "warnings": list(warnings), "errors": []}
        try:
            deck = extract_presentation(ppt, cfg.remove_patterns)
            if w := image_heavy_warning(deck, cfg): rec["warnings"].append(w)
            stem = unique_stem(ppt, root)
            out_json = workspace / "extracted" / f"{stem}.slides.json"
            out_md = workspace / "extracted" / f"{stem}.compact.md"
            out_json.parent.mkdir(parents=True, exist_ok=True)
            out_json.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
            out_md.write_text(deck_to_compact_md(deck, cfg), encoding="utf-8")
            rec.update({"chapter_title": deck["chapter_title"], "slide_count": deck["slide_count"], "intermediate_slides": relpath(out_json, workspace, cfg.absolute_paths), "compact_md": relpath(out_md, workspace, cfg.absolute_paths)})
        except Exception as exc:
            rec["errors"].append(str(exc))
        records.append(rec)
    write_report(workspace, records, cfg)
    return 1 if any(r.get("errors") for r in records) else 0


def handout_json_files(analysis: Path, allow_any_json: bool = False) -> list[Path]:
    if analysis.is_file():
        return [analysis] if allow_any_json or analysis.name.endswith(".handout.json") else []
    if not analysis.exists():
        return []
    return sorted(analysis.glob("*.handout.json"))


def render_cmd(args) -> int:
    warnings: list[str] = []
    cfg = load_config(args.config, warnings)
    analysis = Path(args.analysis).expanduser().resolve(); out = Path(args.output).expanduser().resolve()
    files = handout_json_files(analysis, args.allow_any_json)
    records: list[dict] = []
    if not files:
        records.append({"source_file": str(analysis), "chapter_title": "No handout JSON files found", "warnings": warnings, "errors": ["No *.handout.json files found. Raw slides.json is intentionally refused."]})
        write_report(out, records, cfg); return 1
    for jf in files:
        rec = {"source_file": relpath(jf, out, cfg.absolute_paths), "chapter_title": jf.stem, "warnings": list(warnings), "errors": [], "layout": args.layout}
        try:
            handout = json.loads(jf.read_text(encoding="utf-8"))
            errors = validate_handout_schema(handout)
            if errors: raise ValueError("; ".join(errors))
            stem = safe_name(handout.get("chapter_title", jf.stem))
            docx = out / "docx" / f"{stem}_{safe_name(cfg.output_filename_suffix)}.docx"
            handout_to_docx(handout, docx, layout=args.layout, cfg=cfg)
            rec.update({"chapter_title": handout.get("chapter_title", jf.stem), "slide_count": handout.get("slide_count", 0), "docx": relpath(docx, out, cfg.absolute_paths)})
            if handout.get("generated_by_fallback"): rec["warnings"].append("Generated by deterministic fallback; LLM semantic review is required.")
            if args.export_pdf:
                pdf, warning = export_pdf(docx, out / "pdf")
                if pdf: rec["pdf"] = relpath(pdf, out, cfg.absolute_paths)
                if warning: rec["warnings"].append(warning)
        except Exception as exc:
            rec["errors"].append(str(exc))
        records.append(rec)
    if args.zip_word:
        docx_files = [out / r["docx"] for r in records if r.get("docx")]
        if docx_files:
            zp = out / "word_zip" / cfg.zip_filename; zp.parent.mkdir(parents=True, exist_ok=True)
            with ZipFile(zp, "w", ZIP_DEFLATED) as zf:
                for f in docx_files: zf.write(f, arcname=f.name)
    write_report(out, records, cfg)
    return 1 if any(r.get("errors") for r in records) else 0


def smoke_fallback_handout(deck: dict) -> dict:
    framework, core, tables, image_heavy = [], {}, [], []
    for slide in deck["slides"]:
        title = slide.get("title") or f"Slide {slide['index']}"
        if slide.get("detected_role") == "visual-heavy": image_heavy.append(slide["index"])
        if title not in framework and slide.get("detected_role") not in {"transition-like"}: framework.append(title)
        if slide.get("texts"): core[title] = slide.get("texts", [])[:8]
        for i, table in enumerate(slide.get("tables", []), start=1):
            rows = table.get("rows", [])
            if rows: tables.append({"title": f"{title} table {i}", "headers": rows[0], "rows": rows[1:]})
    return {
        "chapter_title": deck["chapter_title"], "source_file": deck["source_file"],
        "review_goals": ["Smoke-test draft generated from extracted text; replace with caller-authored semantic goals."],
        "knowledge_framework": framework[:20], "core_points": {k: v for k, v in core.items() if v},
        "terms": {}, "comparison_tables": tables[:20], "processes": {},
        "exam_points": ["Caller LLM must replace this fallback field with task-specific review focus."],
        "confusing_points": [], "quick_summary": ["This is a deterministic smoke-test draft, not a final handout."],
        "slide_count": deck["slide_count"], "image_heavy_slides": image_heavy, "generated_by_fallback": True,
    }


def build_cmd(args) -> int:
    workspace = Path(args.output).expanduser().resolve(); workspace.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = ["build is a smoke-test fallback only; high-quality output requires extract -> LLM-authored handout.json -> render."]
    cfg = load_config(args.config, warnings)
    files, unsupported = discover_pptx(Path(args.input).expanduser().resolve(), recursive=args.recursive)
    records = []
    for old in unsupported:
        records.append({"source_file": str(old), "chapter_title": old.stem, "warnings": ["Legacy .ppt is not supported."], "errors": []})
    tmp = workspace / "_fallback_analysis"; tmp.mkdir(parents=True, exist_ok=True)
    for ppt in files:
        try:
            deck = extract_presentation(ppt, cfg.remove_patterns)
            if args.keep_intermediate:
                inter = workspace / "intermediate"; inter.mkdir(parents=True, exist_ok=True)
                stem = unique_stem(ppt, Path(args.input).expanduser().resolve() if Path(args.input).is_dir() else Path(args.input).parent)
                (inter / f"{stem}.slides.json").write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
                (inter / f"{stem}.compact.md").write_text(deck_to_compact_md(deck, cfg), encoding="utf-8")
            h = smoke_fallback_handout(deck)
            (tmp / f"{unique_stem(ppt)}.handout.json").write_text(json.dumps(h, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            records.append({"source_file": str(ppt), "chapter_title": ppt.stem, "warnings": warnings, "errors": [str(exc)]})
    if records and any(r.get("errors") for r in records):
        write_report(workspace, records, cfg); return 1
    ns = argparse.Namespace(analysis=str(tmp), output=str(workspace), export_pdf=args.export_pdf, zip_word=args.zip_word, layout=args.layout, config=args.config, allow_any_json=False)
    return render_cmd(ns)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract PPTX/PPTM files and render caller-authored structured handouts.")
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("extract"); p.add_argument("--input", required=True); p.add_argument("--workspace", required=True); p.add_argument("--config"); p.add_argument("--recursive", action="store_true"); p.set_defaults(func=extract_cmd)
    p = sub.add_parser("render"); p.add_argument("--analysis", required=True); p.add_argument("--output", required=True); p.add_argument("--config"); p.add_argument("--layout", choices=LAYOUT_CHOICES, default="review-margin"); p.add_argument("--export-pdf", action="store_true"); p.add_argument("--zip-word", action="store_true"); p.add_argument("--allow-any-json", action="store_true"); p.set_defaults(func=render_cmd)
    p = sub.add_parser("build"); p.add_argument("--input", required=True); p.add_argument("--output", required=True); p.add_argument("--config"); p.add_argument("--layout", choices=LAYOUT_CHOICES, default="review-margin"); p.add_argument("--export-pdf", action="store_true"); p.add_argument("--zip-word", action="store_true"); p.add_argument("--keep-intermediate", action="store_true"); p.add_argument("--recursive", action="store_true"); p.set_defaults(func=build_cmd)
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
