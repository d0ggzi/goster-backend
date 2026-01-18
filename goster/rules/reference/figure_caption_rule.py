import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


CAPTION_PATTERN = re.compile(r'^Рисунок\s+(\d+(?:\.\d+)?)\s*[–—-]\s*.+', re.IGNORECASE)


class FigureCaptionRule(BaseRule):
    name = "figure_caption_rule"
    description = "Проверка формата подписи рисунка (Рисунок N — Название)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.5.6"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for figure in ctx.model.iter_figures():
            if figure.caption is None:
                yield ValidationError(
                    rule_name=self.name,
                    message="Рисунок не имеет подписи",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=figure.index,
                    paragraph_text="",
                    element_type="figure",
                    current_value="нет подписи",
                    expected_value="Рисунок N — Название",
                    fix_data=None,
                )
                continue

            caption = figure.caption.strip()

            if not CAPTION_PATTERN.match(caption):
                if caption.lower().startswith("рисунок"):
                    yield ValidationError(
                        rule_name=self.name,
                        message="Неверный формат подписи рисунка",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=figure.index,
                        paragraph_text=caption[:100],
                        element_type="figure_caption",
                        current_value=caption[:50],
                        expected_value="Рисунок N — Название",
                        fix_data=None,
                    )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
