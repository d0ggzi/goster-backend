from typing import Iterator

from docx.shared import RGBColor

from ...core.document import DocumentContext
from ...core.model import ParagraphElement, HeadingElement
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class TextColorRule(BaseRule):
    name = "text_color_rule"
    description = "Проверка цвета текста (чёрный)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.1.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for element in ctx.model.elements:
            if not isinstance(element, (ParagraphElement, HeadingElement)):
                continue

            para = element.paragraph
            for j, run in enumerate(para.runs):
                color = run.font.color.rgb
                if color is not None and color != RGBColor(0, 0, 0):
                    yield ValidationError(
                        rule_name=self.name,
                        message="Текст должен быть чёрным",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=element.index,
                        paragraph_text=para.text[:100] if para.text else "",
                        element_type="run",
                        current_value=f"#{color}",
                        expected_value="#000000",
                        fix_data={"element_index": element.index, "run_index": j},
                    )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        run_idx = error.fix_data.get("run_index")

        if element_idx is None or run_idx is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None or run_idx >= len(para.runs):
            return None

        run = para.runs[run_idx]
        before = error.current_value
        run.font.color.rgb = RGBColor(0, 0, 0)

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description="Цвет текста изменён на чёрный",
            before_value=before,
            after_value="#000000",
        )
