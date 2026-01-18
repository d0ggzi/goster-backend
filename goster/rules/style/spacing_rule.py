from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity
from ...gost.requirements import GOSTRequirements
from ...helpers.units import Units


class SpacingRule(BaseRule):
    name = "spacing_rule"
    description = "Проверка межстрочного интервала (1.5)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.1.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        expected_twips = GOSTRequirements.LINE_SPACING_TWIPS
        tolerance = 10

        for element in ctx.model.paragraphs:
            para = element.paragraph
            style_name = para.style.name if para.style else ""

            if style_name.startswith("Heading"):
                continue

            if not para.text or not para.text.strip():
                continue

            pf = para.paragraph_format
            line_spacing = pf.line_spacing

            if line_spacing is None:
                continue

            if hasattr(line_spacing, "twips"):
                spacing_twips = line_spacing.twips
            elif isinstance(line_spacing, float):
                spacing_twips = Units.line_spacing_to_twips(line_spacing)
            else:
                spacing_twips = int(line_spacing)

            if abs(spacing_twips - expected_twips) > tolerance:
                current_lines = Units.twips_to_line_spacing(spacing_twips)
                yield ValidationError(
                    rule_name=self.name,
                    message=f"Неверный межстрочный интервал: {current_lines:.2f}",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=element.index,
                    paragraph_text=para.text[:100] if para.text else "",
                    element_type="paragraph",
                    current_value=f"{current_lines:.2f}",
                    expected_value=f"{GOSTRequirements.LINE_SPACING}",
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
        para.paragraph_format.line_spacing = GOSTRequirements.LINE_SPACING
        after = str(GOSTRequirements.LINE_SPACING)

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=f"Межстрочный интервал изменён на {GOSTRequirements.LINE_SPACING}",
            before_value=before,
            after_value=after,
        )
