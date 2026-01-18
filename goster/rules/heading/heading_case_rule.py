import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class HeadingCaseRule(BaseRule):
    name = "heading_case_rule"
    description = "Проверка регистра заголовков (прописные буквы)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for heading in ctx.model.iter_headings():
            text = heading.text.strip()
            if not text:
                continue

            should_be_uppercase = heading.is_structural or heading.level == 1

            if not should_be_uppercase:
                continue

            text_only = re.sub(r'^[\d.]+\s*', '', text)
            if not text_only:
                continue

            letters = [c for c in text_only if c.isalpha()]
            if not letters:
                continue

            if text_only != text_only.upper():
                yield ValidationError(
                    rule_name=self.name,
                    message="Заголовок раздела должен быть прописными буквами",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=heading.index,
                    paragraph_text=text[:100],
                    element_type="heading",
                    current_value=text_only[:50],
                    expected_value=text_only.upper()[:50],
                    fix_data={"element_index": heading.index},
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        if element_idx is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        before = para.text
        for run in para.runs:
            run.text = run.text.upper()
        after = para.text

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description="Заголовок преобразован в прописные буквы",
            before_value=before[:50],
            after_value=after[:50],
        )
