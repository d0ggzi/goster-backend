from typing import Iterator

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class HeadingNewPageRule(BaseRule):
    name = "heading_new_page_rule"
    description = "Проверка начала раздела с новой страницы"
    gost_reference = "ГОСТ 7.32-2017, п. 6.2.1"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for heading in ctx.model.iter_headings():
            if heading.level > 1 and not heading.is_structural:
                continue

            para = heading.paragraph
            pf = para.paragraph_format

            page_break_before = pf.page_break_before

            if heading.index == 0:
                continue

            if not page_break_before:
                has_page_break = False
                if para.runs:
                    for run in para.runs:
                        xml_str = run._element.xml if isinstance(run._element.xml, str) else str(run._element.xml)
                        if xml_str and 'w:br' in xml_str:
                            if 'w:type="page"' in xml_str:
                                has_page_break = True
                                break

                if not has_page_break:
                    yield ValidationError(
                        rule_name=self.name,
                        message="Раздел должен начинаться с новой страницы",
                        gost_reference=self.gost_reference,
                        severity=Severity.WARNING,
                        paragraph_index=heading.index,
                        paragraph_text=heading.text[:100],
                        element_type="heading",
                        current_value="продолжение страницы",
                        expected_value="новая страница",
                        fix_data={"element_index": heading.index},
                    )

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        element_idx = error.fix_data.get("element_index")
        if element_idx is None:
            return None

        para = ctx.get_paragraph_at(element_idx)
        if para is None:
            return None

        para.paragraph_format.page_break_before = True

        return AppliedFix(
            rule_name=self.name,
            error=error,
            description="Добавлен разрыв страницы перед разделом",
            before_value="продолжение страницы",
            after_value="новая страница",
        )
