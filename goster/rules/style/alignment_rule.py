from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity
from ...gost.requirements import GOSTRequirements
from ...helpers.constants import ALIGNMENT_MAP, ALIGNMENT_NAMES


class AlignmentRule(BaseRule):
    name = "alignment_rule"
    description = "Проверка выравнивания текста (по ширине)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.1.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        expected = ALIGNMENT_MAP[GOSTRequirements.ALIGNMENT_BODY]

        for element in ctx.model.paragraphs:
            para = element.paragraph
            style_name = para.style.name if para.style else ""

            if style_name.startswith("Heading"):
                continue

            alignment = para.alignment
            if alignment is not None and alignment != expected:
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Неверное выравнивание: {ALIGNMENT_NAMES.get(alignment, str(alignment))}",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=element.index,
                    paragraph_text=para.text[:100] if para.text else "",
                    element_type="paragraph",
                    current_value=ALIGNMENT_NAMES.get(alignment, str(alignment)),
                    expected_value=ALIGNMENT_NAMES[expected],
                    fix_data={"element_index": element.index},
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        if element_idx is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        before = error.current_value
        para.alignment = ALIGNMENT_MAP[GOSTRequirements.ALIGNMENT_BODY]
        after = ALIGNMENT_NAMES[ALIGNMENT_MAP[GOSTRequirements.ALIGNMENT_BODY]]

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=f"Выравнивание изменено на '{after}'",
            before_value=before,
            after_value=after,
        )
