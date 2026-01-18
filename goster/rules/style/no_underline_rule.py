from typing import Iterator

from ...core.document import DocumentContext
from ...core.model import ParagraphElement, HeadingElement
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class NoUnderlineRule(BaseRule):
    name = "no_underline_rule"
    description = "Проверка отсутствия подчёркивания"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for element in ctx.model.elements:
            if not isinstance(element, (ParagraphElement, HeadingElement)):
                continue

            para = element.paragraph
            for j, run in enumerate(para.runs):
                if run.underline:
                    yield ValidationError(
                        rule_name=self.name,
                        message="Подчёркивание текста запрещено",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=element.index,
                        paragraph_text=run.text[:50] if run.text else "",
                        element_type="run",
                        current_value="подчёркнуто",
                        expected_value="без подчёркивания",
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
        run.underline = False

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description="Подчёркивание удалено",
            before_value="подчёркнуто",
            after_value="без подчёркивания",
        )
