from __future__ import annotations

from enum import Enum
from typing import TYPE_CHECKING

from docx.shared import RGBColor
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    from .document import DocumentContext
    from .report import ValidationError

from .report import AppliedHighlight


class HighlightColor(Enum):
    RED = RGBColor(255, 0, 0)
    ORANGE = RGBColor(255, 165, 0)
    YELLOW = RGBColor(255, 255, 0)
    BLUE = RGBColor(0, 100, 255)


class DocumentHighlighter:
    def __init__(self, ctx: "DocumentContext"):
        self.ctx = ctx
        self.highlights: list[AppliedHighlight] = []

    def highlight_paragraph(
        self,
        para: Paragraph,
        color: HighlightColor = HighlightColor.RED,
    ) -> bool:
        if not para.runs:
            return False

        for run in para.runs:
            run.font.color.rgb = color.value

        return True

    def highlight_run(
        self,
        para: Paragraph,
        run_index: int,
        color: HighlightColor = HighlightColor.RED,
    ) -> bool:
        if run_index >= len(para.runs):
            return False

        para.runs[run_index].font.color.rgb = color.value
        return True

    def highlight_error(
        self,
        error: "ValidationError",
        color: HighlightColor | None = None,
    ) -> AppliedHighlight | None:
        from .report import Severity

        if error.paragraph_index is None:
            return None

        para = self.ctx.get_paragraph_at(error.paragraph_index)
        if para is None:
            return None

        if color is None:
            color = {
                Severity.ERROR: HighlightColor.RED,
                Severity.WARNING: HighlightColor.ORANGE,
                Severity.INFO: HighlightColor.BLUE,
            }.get(error.severity, HighlightColor.RED)

        run_index = error.fix_data.get("run_index") if error.fix_data else None

        if run_index is not None:
            success = self.highlight_run(para, run_index, color)
        else:
            success = self.highlight_paragraph(para, color)

        if not success:
            return None

        highlight = AppliedHighlight(
            rule_name=error.rule_name,
            error=error,
            paragraph_index=error.paragraph_index,
            color=color.name.lower(),
            description=f"Подсвечено: {error.message}",
        )
        self.highlights.append(highlight)
        return highlight

    def highlight_all(
        self,
        errors: list["ValidationError"],
    ) -> list[AppliedHighlight]:
        results = []
        for error in errors:
            highlight = self.highlight_error(error)
            if highlight:
                results.append(highlight)
        return results
