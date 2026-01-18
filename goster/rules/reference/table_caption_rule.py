import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


CAPTION_PATTERN = re.compile(r'^Таблица\s+(\d+(?:\.\d+)?)\s*[–—-]\s*.+', re.IGNORECASE)


class TableCaptionRule(BaseRule):
    name = "table_caption_rule"
    description = "Проверка формата подписи таблицы (Таблица N — Название)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.5.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for table in ctx.model.iter_tables():
            if table.caption is None:
                yield ValidationError(
                    rule_name=self.name,
                    message="Таблица не имеет подписи",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=table.index,
                    paragraph_text="",
                    element_type="table",
                    current_value="нет подписи",
                    expected_value="Таблица N — Название",
                    fix_data=None,
                )
                continue

            caption = table.caption.strip()

            if not CAPTION_PATTERN.match(caption):
                if caption.lower().startswith("таблица"):
                    yield ValidationError(
                        rule_name=self.name,
                        message="Неверный формат подписи таблицы",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=table.caption_paragraph_index,
                        paragraph_text=caption[:100],
                        element_type="table_caption",
                        current_value=caption[:50],
                        expected_value="Таблица N — Название",
                        fix_data=None,
                    )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
