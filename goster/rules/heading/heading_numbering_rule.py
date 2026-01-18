import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


CHAPTER_PATTERN = re.compile(r"^ГЛАВА\s+(\d+)", re.IGNORECASE)
NUMBER_PATTERN = re.compile(r"^(\d+(?:\.\d+)*)")


class HeadingNumberingRule(BaseRule):
    name = "heading_numbering_rule"
    description = "Проверка нумерации заголовков"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        expected_counters = {}

        for heading in ctx.model.iter_headings():
            if heading.is_structural:
                continue

            text = heading.text.strip()
            if not text:
                continue

            number = None
            chapter_match = CHAPTER_PATTERN.match(text)
            if chapter_match:
                number = chapter_match.group(1)
            else:
                num_match = NUMBER_PATTERN.match(text)
                if num_match:
                    number = num_match.group(1)

            if not number:
                yield ValidationError(
                    rule_name=self.name,
                    message="Заголовок раздела должен быть пронумерован",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=heading.index,
                    paragraph_text=text[:100],
                    element_type="heading",
                    current_value="отсутствует",
                    expected_value="нумерация (например, 1, 1.1, 2.1.3)",
                    fixable=False,
                )
                continue

            try:
                parts = [int(p) for p in number.split(".")]
            except ValueError:
                continue
            level = len(parts)

            if level not in expected_counters:
                for i in range(1, level + 1):
                    if i not in expected_counters:
                        expected_counters[i] = 0

            expected_counters[level] += 1
            for deeper in list(expected_counters.keys()):
                if deeper > level:
                    expected_counters[deeper] = 0

            expected_parts = [expected_counters.get(i, 0) for i in range(1, level + 1)]
            expected_number = ".".join(str(p) for p in expected_parts)

            if number != expected_number:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Неверная нумерация заголовка: {number}",
                    gost_reference=self.gost_reference,
                    severity=Severity.WARNING,
                    paragraph_index=heading.index,
                    paragraph_text=text[:100],
                    element_type="heading",
                    current_value=number,
                    expected_value=expected_number,
                    fixable=False,
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
