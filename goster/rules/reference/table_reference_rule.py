from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class TableReferenceRule(BaseRule):
    name = "table_reference_rule"
    description = "Проверка ссылок на таблицы (должна быть ссылка до таблицы)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.5.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for table in ctx.model.iter_tables():
            if not table.number:
                if table.caption_paragraph_index is not None:
                    yield ValidationError(
                        rule_name=self.name,
                        message="Таблица не имеет номера",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=table.caption_paragraph_index,
                        paragraph_text=table.caption[:100] if table.caption else "",
                        element_type="table",
                        current_value="без номера",
                        expected_value="Таблица N — Название",
                        fix_data=None,
                    )
                continue

            if not table.has_reference_before:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"На таблицу {table.number} нет ссылки в тексте до её появления",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=table.caption_paragraph_index,
                    paragraph_text=table.caption[:100] if table.caption else "",
                    element_type="table",
                    current_value="нет ссылки",
                    expected_value=f"ссылка на таблицу {table.number} в тексте",
                    fix_data=None,
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
