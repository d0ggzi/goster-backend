from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


REQUIRED_SECTIONS = [
    ("ВВЕДЕНИЕ", ["ВВЕДЕНИЕ"]),
    ("ЗАКЛЮЧЕНИЕ", ["ЗАКЛЮЧЕНИЕ"]),
    ("СПИСОК ИСТОЧНИКОВ", ["СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ЛИТЕРАТУРЫ", "БИБЛИОГРАФИЧЕСКИЙ СПИСОК"]),
]

OPTIONAL_SECTIONS = [
    ("РЕФЕРАТ", ["РЕФЕРАТ"]),
    ("СОДЕРЖАНИЕ", ["СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"]),
    ("СПИСОК СОКРАЩЕНИЙ", ["СПИСОК СОКРАЩЕНИЙ", "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ", "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ"]),
    ("ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ", ["ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ"]),
    ("ПРИЛОЖЕНИЯ", ["ПРИЛОЖЕНИЕ"]),
]


class RequiredSectionsRule(BaseRule):
    name = "required_sections_rule"
    description = "Проверка наличия обязательных разделов"
    gost_reference = "ГОСТ 7.32-2017, п. 5.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        found_sections = set()

        for heading in ctx.model.iter_headings():
            text_upper = heading.text.strip().upper()

            for name, variants in REQUIRED_SECTIONS + OPTIONAL_SECTIONS:
                if any(text_upper.startswith(v) or text_upper == v for v in variants):
                    found_sections.add(name)
                    break

        for name, variants in REQUIRED_SECTIONS:
            if name not in found_sections:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Отсутствует обязательный раздел: {name}",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=None,
                    paragraph_text="",
                    element_type="document",
                    current_value="отсутствует",
                    expected_value=name,
                    fixable=False,
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
