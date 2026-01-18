import re
from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class SubheadingCaseRule(BaseRule):
    name = "subheading_case_rule"
    description = "Проверка регистра подзаголовков (только первая буква прописная)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2.3"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for heading in ctx.model.iter_headings():
            if heading.level < 2:
                continue

            if heading.is_structural:
                continue

            text = heading.text.strip()
            if not text:
                continue

            text_only = re.sub(r'^[\d.]+\s*', '', text)
            if not text_only:
                continue

            if text_only == text_only.upper():
                yield ValidationError(
                    rule_name=self.name,
                    message="Подзаголовок не должен быть полностью прописными",
                    gost_reference=self.gost_reference,
                    severity=Severity.WARNING,
                    paragraph_index=heading.index,
                    paragraph_text=text[:100],
                    element_type="heading",
                    current_value=text_only[:30],
                    expected_value="Только первая буква прописная",
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
            if run.text:
                run.text = run.text.capitalize()
        after = para.text

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description="Регистр подзаголовка исправлен",
            before_value=before[:30],
            after_value=after[:30],
        )
