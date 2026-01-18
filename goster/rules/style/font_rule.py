from typing import Iterator

from docx.shared import Pt

from ...core.document import DocumentContext
from ...core.model import ParagraphElement, HeadingElement
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity
from ...gost.requirements import GOSTRequirements


class FontRule(BaseRule):
    name = "font_rule"
    description = "Проверка шрифта (Times New Roman, 14 пт)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.1.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for element in ctx.model.elements:
            if not isinstance(element, (ParagraphElement, HeadingElement)):
                continue

            para = element.paragraph
            for j, run in enumerate(para.runs):
                font = run.font

                if font.name and font.name != GOSTRequirements.FONT_NAME:
                    yield ValidationError(
                        rule_name=self.name,
                        message=f"Неверный шрифт: {font.name}",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=element.index,
                        paragraph_text=para.text[:100] if para.text else "",
                        element_type="run",
                        current_value=font.name,
                        expected_value=GOSTRequirements.FONT_NAME,
                        fix_data={
                            "element_index": element.index,
                            "run_index": j,
                            "error_type": "wrong_name",
                        },
                    )

                if font.size:
                    size_pt = font.size.pt
                    if not (GOSTRequirements.FONT_SIZE_MIN_PT <= size_pt <= GOSTRequirements.FONT_SIZE_MAX_PT):
                        yield ValidationError(
                            rule_name=self.name,
                            message=f"Неверный размер шрифта: {size_pt} пт",
                            gost_reference=self.gost_reference,
                            severity=Severity.ERROR,
                            paragraph_index=element.index,
                            paragraph_text=para.text[:100] if para.text else "",
                            element_type="run",
                            current_value=f"{size_pt} пт",
                            expected_value=f"{GOSTRequirements.FONT_SIZE_MIN_PT}-{GOSTRequirements.FONT_SIZE_MAX_PT} пт",
                            fix_data={
                                "element_index": element.index,
                                "run_index": j,
                                "error_type": "wrong_size",
                            },
                        )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        run_idx = error.fix_data.get("run_index")
        error_type = error.fix_data.get("error_type")

        if element_idx is None or run_idx is None or error_type is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        if run_idx >= len(para.runs):
            return None

        run = para.runs[run_idx]
        before = error.current_value

        if error_type == "wrong_size":
            run.font.size = Pt(GOSTRequirements.FONT_SIZE_PT)
            after = f"{GOSTRequirements.FONT_SIZE_PT} пт"
            desc = f"Размер шрифта изменён на {GOSTRequirements.FONT_SIZE_PT} пт"
        else:
            run.font.name = GOSTRequirements.FONT_NAME
            after = GOSTRequirements.FONT_NAME
            desc = f"Шрифт изменён на {GOSTRequirements.FONT_NAME}"

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=desc,
            before_value=before,
            after_value=after,
        )
