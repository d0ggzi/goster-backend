import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class ListFormatRule(BaseRule):
    name = "list_format_rule"
    description = "Проверка формата списков"
    gost_reference = "ГОСТ 7.32-2017, п. 6.4"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for para_element in ctx.model.iter_paragraphs():
            para = para_element.paragraph
            text = para.text.strip()

            if not text:
                continue

            if text.startswith('•') or text.startswith('●') or text.startswith('○'):
                yield ValidationError(
                    rule_name=self.name,
                    message="Маркер списка должен быть дефисом (–)",
                    gost_reference=self.gost_reference,
                    severity=Severity.WARNING,
                    paragraph_index=para_element.index,
                    paragraph_text=text[:100],
                    element_type="list_item",
                    current_value=text[0],
                    expected_value="–",
                    fix_data={"element_index": para_element.index},
                )

            upper_match = re.match(r'^[А-ЯA-Z]\)', text)
            if upper_match:
                yield ValidationError(
                    rule_name=self.name,
                    message="Буквенный маркер списка должен быть строчной буквой",
                    gost_reference=self.gost_reference,
                    severity=Severity.WARNING,
                    paragraph_index=para_element.index,
                    paragraph_text=text[:100],
                    element_type="list_item",
                    current_value="прописная буква",
                    expected_value="строчная буква",
                    fix_data=None,
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index") if error.fix_data else None
        if element_idx is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None or not para.runs:
            return None

        first_run = para.runs[0]
        before = first_run.text[0] if first_run.text else ""

        if first_run.text and first_run.text[0] in ['•', '●', '○']:
            first_run.text = '–' + first_run.text[1:]

            return AppliedFix(
                rule_name=self.name,
                error=error,
                description="Маркер списка изменён на дефис",
                before_value=before,
                after_value="–",
            )

        return None
