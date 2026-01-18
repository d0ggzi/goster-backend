import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


APPENDIX_PATTERN = re.compile(r"^ПРИЛОЖЕНИЕ\s+([А-Я])", re.IGNORECASE)
INVALID_APPENDIX_LETTERS = frozenset(["Ё", "Й", "О", "Ч", "Ь", "Ы", "Ъ"])
VALID_APPENDIX_ORDER = "АБВГДЕЖЗИКЛМНПРСТУФХЦШЩЭЮЯ"


class AppendixRule(BaseRule):
    name = "appendix_rule"
    description = "Проверка оформления приложений"
    gost_reference = "ГОСТ 7.32-2017, п. 6.14"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        found_appendices = []

        for heading in ctx.model.iter_headings():
            text = heading.text.strip().upper()

            match = APPENDIX_PATTERN.match(text)
            if not match:
                continue

            letter = match.group(1).upper()
            found_appendices.append((heading, letter))

            if letter in INVALID_APPENDIX_LETTERS:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Буква '{letter}' не используется для обозначения приложений",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=heading.index,
                    paragraph_text=heading.text[:100],
                    element_type="appendix",
                    current_value=letter,
                    expected_value="А, Б, В, Г...",
                    fix_data=None,
                )

        for i, (heading, letter) in enumerate(found_appendices):
            if letter in INVALID_APPENDIX_LETTERS:
                continue

            expected_letter = (
                VALID_APPENDIX_ORDER[i] if i < len(VALID_APPENDIX_ORDER) else None
            )
            if expected_letter and letter != expected_letter:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Неверный порядок приложений: '{letter}' вместо '{expected_letter}'",
                    gost_reference=self.gost_reference,
                    severity=Severity.WARNING,
                    paragraph_index=heading.index,
                    paragraph_text=heading.text[:100],
                    element_type="appendix",
                    current_value=letter,
                    expected_value=expected_letter,
                    fix_data=None,
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
