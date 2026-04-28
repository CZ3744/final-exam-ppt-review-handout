from pathlib import Path

from ppt_review_handout.cli_generic import chapter_index, chinese_to_int, discover_pptx, safe_name


def test_chinese_to_int():
    assert chinese_to_int("一") == 1
    assert chinese_to_int("十") == 10
    assert chinese_to_int("十一") == 11
    assert chinese_to_int("二十三") == 23
    assert chinese_to_int("101") == 101


def test_chapter_index():
    assert chapter_index("chapter 3 materials.pptx") == 3
    assert chapter_index("06 final.pptx") == 6


def test_safe_name():
    assert safe_name('A: B/C') == "A_ B_C"


def test_discover_supported_and_unsupported(tmp_path: Path):
    (tmp_path / "02 B.pptx").write_text("", encoding="utf-8")
    (tmp_path / "01 A.pptm").write_text("", encoding="utf-8")
    (tmp_path / "old.ppt").write_text("", encoding="utf-8")
    files, unsupported = discover_pptx(tmp_path)
    assert [p.name for p in files] == ["01 A.pptm", "02 B.pptx"]
    assert [p.name for p in unsupported] == ["old.ppt"]


def test_discover_missing_path_is_tolerant(tmp_path: Path):
    files, unsupported = discover_pptx(tmp_path / "missing")
    assert files == []
    assert unsupported == []
