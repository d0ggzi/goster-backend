from typing import Iterator

from docx.enum.text import WD_ALIGN_PARAGRAPH

from ...core.document import DocumentContext
from ...core.rule import BaseRule
from ...core.report import ValidationError, AppliedFix, Severity


class PageNumberingRule(BaseRule):
    name = "page_numbering_rule"
    description = "Проверка нумерации страниц (внизу по центру)"
    gost_reference = "ГОСТ 7.32-2017, п. 6.3"

    def validate(self, ctx: DocumentContext) -> Iterator[ValidationError]:
        for i, section in enumerate(ctx.document.sections):
            footer = section.footer

            if not footer.is_linked_to_previous:
                has_page_number = False
                for para in footer.paragraphs:
                    xml_str = para._element.xml if isinstance(para._element.xml, str) else str(para._element.xml)
                    if 'w:fldChar' in xml_str or 'PAGE' in xml_str:
                        has_page_number = True
                        if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                            yield ValidationError(
                                rule_name=self.name,
                                message="Номер страницы должен быть по центру",
                                gost_reference=self.gost_reference,
                                severity=Severity.WARNING,
                                paragraph_index=None,
                                paragraph_text=f"Секция {i + 1}",
                                element_type="footer",
                                current_value="не по центру",
                                expected_value="по центру",
                                fix_data=None,
                            )
                        break

    def fix(self, ctx: DocumentContext, error: ValidationError) -> AppliedFix | None:
        return None
