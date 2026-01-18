from typing import Iterator

from docx.shared import Pt

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class ParagraphSpacingRule(BaseRule):
    name = "paragraph_spacing_rule"
    description = "Проверка интервалов до/после абзацев"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2.2"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for heading in ctx.model.iter_headings():
            para = heading.paragraph
            pf = para.paragraph_format

            if pf.space_before and pf.space_before.pt > 24:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Слишком большой интервал перед заголовком: {pf.space_before.pt} пт",
                    gost_reference=self.gost_reference,
                    severity=Severity.WARNING,
                    paragraph_index=heading.index,
                    paragraph_text=heading.text[:100],
                    element_type="heading",
                    current_value=f"{pf.space_before.pt} пт",
                    expected_value="до 24 пт",
                    fix_data={"element_index": heading.index, "type": "before"},
                )

            if pf.space_after and pf.space_after.pt > 18:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Слишком большой интервал после заголовка: {pf.space_after.pt} пт",
                    gost_reference=self.gost_reference,
                    severity=Severity.WARNING,
                    paragraph_index=heading.index,
                    paragraph_text=heading.text[:100],
                    element_type="heading",
                    current_value=f"{pf.space_after.pt} пт",
                    expected_value="до 18 пт",
                    fix_data={"element_index": heading.index, "type": "after"},
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        fix_type = error.fix_data.get("type")

        if element_idx is None or fix_type is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        before = error.current_value

        if fix_type == "before":
            para.paragraph_format.space_before = Pt(12)
            after = "12 пт"
        else:
            para.paragraph_format.space_after = Pt(12)
            after = "12 пт"

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=f"Интервал {'перед' if fix_type == 'before' else 'после'} заголовком изменён",
            before_value=before,
            after_value=after,
        )
