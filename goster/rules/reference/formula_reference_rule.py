from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class FormulaReferenceRule(BaseRule):
    name = "formula_reference_rule"
    description = "Проверка нумерации формул"
    gost_reference = "ГОСТ 7.32-2017, п. 6.4"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for formula in ctx.model.formulas:
            if not formula.number:
                yield ValidationError(
                    rule_name=self.name,
                    message="Формула должна быть пронумерована",
                    gost_reference=self.gost_reference,
                    severity=Severity.ERROR,
                    paragraph_index=formula.index,
                    paragraph_text=formula.text[:100] if formula.text else "",
                    element_type="formula",
                    current_value="без номера",
                    expected_value="(N) или (N.M)",
                    fix_data=None,
                )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
