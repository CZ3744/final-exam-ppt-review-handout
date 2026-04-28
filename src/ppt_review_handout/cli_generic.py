from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE_TYPE

from . import workflow_cli as _workflow


def _enum_member(name: str):
    return getattr(MSO_SHAPE_TYPE, name, None)


def visual_weight(shape) -> int:
    """Version-tolerant visual element detector.

    python-pptx has changed/varied enum names across versions. Avoid direct
    access to optional members such as MEDIA at import/runtime sites that would
    otherwise raise AttributeError on older installations.
    """
    shape_type = getattr(shape, "shape_type", None)
    if shape_type == _enum_member("PICTURE"):
        return 1
    if getattr(shape, "has_chart", False):
        return 1
    if getattr(shape, "has_table", False):
        return 1
    optional_visual_types = {
        value
        for value in (
            _enum_member("MEDIA"),
            _enum_member("MOVIE"),
            _enum_member("OLE_OBJECT"),
            _enum_member("PLACEHOLDER"),
        )
        if value is not None
    }
    return 1 if shape_type in optional_visual_types else 0


# Patch the active implementation before exposing its CLI.
_workflow.visual_weight = visual_weight

SkillConfig = _workflow.SkillConfig
chinese_to_int = _workflow.chinese_to_int
chapter_index = _workflow.chapter_index
safe_name = _workflow.safe_name
discover_pptx = _workflow.discover_pptx
validate_handout_schema = _workflow.validate_handout_schema
handout_to_docx = _workflow.handout_to_docx
main = _workflow.main


if __name__ == "__main__":
    raise SystemExit(main())
