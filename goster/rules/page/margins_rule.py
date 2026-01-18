from typing import Iterator

from docx.shared import Twips

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity
from ...gost.requirements import GOSTRequirements
from ...helpers.units import Units


class MarginsRule(BaseRule):
    name = "margins_rule"
    description = "Проверка полей страницы"
    gost_reference = "ГОСТ 7.32-2017, п. 6.1.2"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        tolerance = Units.cm_to_twips(0.1)

        for i, section in enumerate(ctx.document.sections):
            for name, current, check_fn, expected_str in [
                (
                    "левое",
                    section.left_margin,
                    lambda t: abs(t - GOSTRequirements.MARGIN_LEFT_TWIPS) <= tolerance,
                    f"{GOSTRequirements.MARGIN_LEFT_CM} см",
                ),
                (
                    "правое",
                    section.right_margin,
                    lambda t: GOSTRequirements.MARGIN_RIGHT_MIN_TWIPS - tolerance
                    <= t
                    <= GOSTRequirements.MARGIN_RIGHT_MAX_TWIPS + tolerance,
                    f"{GOSTRequirements.MARGIN_RIGHT_MIN_CM}-{GOSTRequirements.MARGIN_RIGHT_MAX_CM} см",
                ),
                (
                    "верхнее",
                    section.top_margin,
                    lambda t: abs(t - GOSTRequirements.MARGIN_TOP_TWIPS) <= tolerance,
                    f"{GOSTRequirements.MARGIN_TOP_CM} см",
                ),
                (
                    "нижнее",
                    section.bottom_margin,
                    lambda t: abs(t - GOSTRequirements.MARGIN_BOTTOM_TWIPS)
                    <= tolerance,
                    f"{GOSTRequirements.MARGIN_BOTTOM_CM} см",
                ),
            ]:
                if current is None:
                    continue

                current_twips = (
                    current.twips if hasattr(current, "twips") else int(current)
                )

                if not check_fn(current_twips):
                    current_cm = Units.twips_to_cm(current_twips)
                    yield ValidationError(
                        rule_name=self.name,
                        message=f"Неверное {name} поле: {current_cm:.1f} см",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=None,
                        paragraph_text=f"Секция {i + 1}",
                        element_type="section",
                        current_value=f"{current_cm:.1f} см",
                        expected_value=expected_str,
                        fix_data={"section_index": i, "margin_name": name},
                    )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        section_idx = error.fix_data.get("section_index")
        margin_name = error.fix_data.get("margin_name")

        if section_idx is None or margin_name is None:
            return None

        if section_idx >= len(ctx.document.sections):
            return None

        section = ctx.document.sections[section_idx]
        before = error.current_value

        margin_map = {
            "левое": (
                "left_margin",
                GOSTRequirements.MARGIN_LEFT_TWIPS,
                GOSTRequirements.MARGIN_LEFT_CM,
            ),
            "правое": (
                "right_margin",
                GOSTRequirements.MARGIN_RIGHT_MIN_TWIPS,
                GOSTRequirements.MARGIN_RIGHT_MIN_CM,
            ),
            "верхнее": (
                "top_margin",
                GOSTRequirements.MARGIN_TOP_TWIPS,
                GOSTRequirements.MARGIN_TOP_CM,
            ),
            "нижнее": (
                "bottom_margin",
                GOSTRequirements.MARGIN_BOTTOM_TWIPS,
                GOSTRequirements.MARGIN_BOTTOM_CM,
            ),
        }

        attr, twips, cm = margin_map[margin_name]
        setattr(section, attr, Twips(twips))
        after = f"{cm} см"

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=f"{margin_name.capitalize()} поле изменено на {cm} см",
            before_value=before,
            after_value=after,
        )
