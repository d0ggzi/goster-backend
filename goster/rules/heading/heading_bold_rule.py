from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class HeadingBoldRule(BaseRule):
    name = "heading_bold_rule"
    description = "Проверка начертания заголовков"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for heading in ctx.model.iter_headings():
            para = heading.paragraph
            if not para.runs:
                continue

            runs_with_text = [run for run in para.runs if run.text.strip()]
            if not runs_with_text:
                continue

            all_bold = all(run.bold for run in runs_with_text)
            all_italic = all(run.italic for run in runs_with_text)

            if heading.level == 1 or heading.is_structural:
                if not all_bold:
                    yield ValidationError(
                        rule_name=self.name,
                        message="Заголовок раздела должен быть полужирным",
                        gost_reference=self.gost_reference,
                        severity=Severity.ERROR,
                        paragraph_index=heading.index,
                        paragraph_text=heading.text[:100],
                        element_type="heading",
                        current_value="не полужирный",
                        expected_value="полужирный",
                        fix_data={
                            "element_index": heading.index,
                            "action": "make_bold",
                        },
                    )
            elif heading.level == 2:
                if all_bold:
                    yield ValidationError(
                        rule_name=self.name,
                        message="Подзаголовок не должен быть полужирным",
                        gost_reference="ГОСТ 7.32-2017, п. 6.2.3",
                        severity=Severity.WARNING,
                        paragraph_index=heading.index,
                        paragraph_text=heading.text[:100],
                        element_type="heading",
                        current_value="полужирный",
                        expected_value="обычный",
                        fix_data={
                            "element_index": heading.index,
                            "action": "remove_bold",
                        },
                    )
            elif heading.level >= 3:
                if not all_italic:
                    yield ValidationError(
                        rule_name=self.name,
                        message="Пункт/подпункт должен быть курсивом",
                        gost_reference="ГОСТ 7.32-2017, п. 6.2.4",
                        severity=Severity.WARNING,
                        paragraph_index=heading.index,
                        paragraph_text=heading.text[:100],
                        element_type="heading",
                        current_value="не курсив",
                        expected_value="курсив",
                        fix_data={
                            "element_index": heading.index,
                            "action": "make_italic",
                        },
                    )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        action = error.fix_data.get("action")
        if element_idx is None or action is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        before = error.current_value

        if action == "make_bold":
            for run in para.runs:
                run.bold = True
            after = "полужирный"
            desc = "Заголовок сделан полужирным"
        elif action == "remove_bold":
            for run in para.runs:
                run.bold = False
            after = "обычный"
            desc = "Полужирное начертание убрано"
        elif action == "make_italic":
            for run in para.runs:
                run.italic = True
            after = "курсив"
            desc = "Добавлен курсив"
        else:
            return None

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description=desc,
            before_value=before,
            after_value=after,
        )
