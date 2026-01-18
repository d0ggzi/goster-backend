from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class HeadingPunctuationRule(BaseRule):
    name = "heading_punctuation_rule"
    description = "Проверка отсутствия точки в конце заголовка"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for heading in ctx.model.iter_headings():
            text = heading.text.strip()
            if not text:
                continue

            if text.endswith('.'):
                yield ValidationError(
                    rule_name=self.name,
                    message="Заголовок не должен заканчиваться точкой",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=heading.index,
                    paragraph_text=text[:100],
                    element_type="heading",
                    current_value=text[-20:] if len(text) > 20 else text,
                    expected_value="без точки в конце",
                    fix_data={"element_index": heading.index},
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        if element_idx is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        if para.runs:
            last_run = para.runs[-1]
            if last_run.text.endswith('.'):
                before = last_run.text
                last_run.text = last_run.text.rstrip('.')
                after = last_run.text

                return AppliedFix(
                    rule_name=self.name,
                    error=error,
                    description="Удалена точка в конце заголовка",
                    before_value=before[-20:] if len(before) > 20 else before,
                    after_value=after[-20:] if len(after) > 20 else after,
                )

        return None
