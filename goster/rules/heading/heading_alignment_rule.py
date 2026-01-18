from typing import Iterator

from docx.enum.text import WD_ALIGN_PARAGRAPH

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity
from ...helpers.constants import ALIGNMENT_NAMES


class HeadingAlignmentRule(BaseRule):
    name = "heading_alignment_rule"
    description = "Проверка выравнивания заголовков"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for heading in ctx.model.iter_headings():
            para = heading.paragraph
            alignment = para.alignment

            if heading.level == 1 or heading.is_structural:
                expected = WD_ALIGN_PARAGRAPH.CENTER
                expected_name = "по центру"
            else:
                expected = WD_ALIGN_PARAGRAPH.LEFT
                expected_name = "по левому краю (с абзацного отступа)"

            if alignment is not None and alignment != expected:
                yield ValidationError(
                    rule_name=self.name,
                    message="Неверное выравнивание заголовка",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR if heading.level == 1 else Severity.WARNING,
                    paragraph_index=heading.index,
                    paragraph_text=heading.text[:100],
                    element_type="heading",
                    current_value=ALIGNMENT_NAMES.get(alignment, str(alignment)),
                    expected_value=expected_name,
                    fix_data={"element_index": heading.index, "expected": expected},
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        expected = error.fix_data.get("expected")
        if element_idx is None or expected is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        before = error.current_value
        para.alignment = expected
        after = ALIGNMENT_NAMES.get(expected, str(expected))

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=f"Выравнивание заголовка изменено на '{after}'",
            before_value=before,
            after_value=after,
        )
