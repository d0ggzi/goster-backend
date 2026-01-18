from typing import Iterator

from docx.shared import Twips

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity
from ...gost.requirements import GOSTRequirements
from ...helpers.units import Units


class IndentRule(BaseRule):
    name = "indent_rule"
    description = "Проверка абзацного отступа (1.25 см)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.1.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        expected_twips = GOSTRequirements.FIRST_LINE_INDENT_TWIPS
        tolerance = Units.cm_to_twips(0.05)

        for element in ctx.model.paragraphs:
            para = element.paragraph
            style_name = para.style.name if para.style else ""

            if style_name.startswith("Heading"):
                continue

            if not para.text or not para.text.strip():
                continue

            pf = para.paragraph_format
            first_line = pf.first_line_indent

            if first_line is None:
                continue

            first_line_twips = first_line.twips if hasattr(first_line, "twips") else int(first_line)

            if abs(first_line_twips - expected_twips) > tolerance:
                current_cm = Units.twips_to_cm(first_line_twips)
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Неверный абзацный отступ: {current_cm:.2f} см",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=element.index,
                    paragraph_text=para.text[:100] if para.text else "",
                    element_type="paragraph",
                    current_value=f"{current_cm:.2f} см",
                    expected_value=f"{GOSTRequirements.FIRST_LINE_INDENT_CM} см",
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
        para.paragraph_format.first_line_indent = Twips(GOSTRequirements.FIRST_LINE_INDENT_TWIPS)
        after = f"{GOSTRequirements.FIRST_LINE_INDENT_CM} см"

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=f"Абзацный отступ изменён на {GOSTRequirements.FIRST_LINE_INDENT_CM} см",
            before_value=before,
            after_value=after,
        )
