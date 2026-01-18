from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class FigureReferenceRule(BaseRule):
    name = "figure_reference_rule"
    description = "Проверка ссылок на рисунки (должна быть ссылка до рисунка)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.5.6"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for figure in ctx.model.iter_figures():
            if not figure.number:
                yield ValidationError(
                    rule_name=self.name,
                    message="Рисунок не имеет номера",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=figure.index,
                    paragraph_text=figure.caption[:100] if figure.caption else "",
                    element_type="figure",
                    current_value="без номера",
                    expected_value="Рисунок N — Название",
                    fix_data=None,
                )
                continue

            if not figure.has_reference_before:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"На рисунок {figure.number} нет ссылки в тексте до его появления",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=figure.index,
                    paragraph_text=figure.caption[:100] if figure.caption else "",
                    element_type="figure",
                    current_value="нет ссылки",
                    expected_value=f"ссылка на рисунок {figure.number} в тексте",
                    fix_data=None,
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
