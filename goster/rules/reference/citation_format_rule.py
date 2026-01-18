import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


INVALID_CITATION_PATTERNS = [
    (re.compile(r'(?<!\S)\(\d+\)(?!\s*$)'), "круглые скобки вместо квадратных"),
    (re.compile(r'\[\d+\.\]'), "точка после номера"),
]


class CitationFormatRule(BaseRule):
    name = "citation_format_rule"
    description = "Проверка формата ссылок на источники [N]"
    gost_reference = "ГОСТ 7.32-2017, п. 6.8"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for para_element in ctx.model.iter_paragraphs():
            text = para_element.text
            if not text:
                continue

            for pattern, issue in INVALID_CITATION_PATTERNS:
                for match in pattern.finditer(text):
                    yield ValidationError(
                        rule_name=self.name,
                        message=f"Неверный формат ссылки: {issue}",
                        gost_reference=self.gost_reference,
                        severity=Severity.WARNING,
                        paragraph_index=para_element.index,
                        paragraph_text=text[:100],
                        element_type="citation",
                        current_value=match.group(),
                        expected_value="[N] или [N, с. XX]",
                        fix_data=None,
                    )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
