from zipfile import ZipFile
from lxml import etree
import shutil


def get_document(input_file):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(input_file) as docx_zip:
        xml_content = docx_zip.read("word/document.xml")
    root = etree.fromstring(xml_content)
    return root


def save_docx(input_docx, output_docx,
              new_styles_root=None,
              new_document_root=None,
              new_numbering_root=None):
    new_styles_xml = (
        etree.tostring(new_styles_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        if new_styles_root is not None else None
    )

    new_document_xml = (
        etree.tostring(new_document_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        if new_document_root is not None else None
    )

    new_numbering_xml = (
        etree.tostring(new_numbering_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        if new_numbering_root is not None else None
    )

    temp_zip_path = output_docx + "_temp.zip"

    with ZipFile(input_docx, 'r') as zin, ZipFile(temp_zip_path, 'w') as zout:

        has_numbering_in_input = False

        for item in zin.infolist():

            if item.filename == 'word/styles.xml' and new_styles_xml is not None:
                zout.writestr('word/styles.xml', new_styles_xml)

            elif item.filename == 'word/document.xml' and new_document_xml is not None:
                zout.writestr('word/document.xml', new_document_xml)

            elif item.filename == 'word/numbering.xml':
                has_numbering_in_input = True
                if new_numbering_xml is not None:
                    zout.writestr('word/numbering.xml', new_numbering_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

            else:
                zout.writestr(item, zin.read(item.filename))

        if not has_numbering_in_input and new_numbering_xml is not None:
            zout.writestr('word/numbering.xml', new_numbering_xml)

    shutil.move(temp_zip_path, output_docx)
    print(f"✅ Файл успешно сохранён: {output_docx}")


def get_table_of_contents(document):
    import re
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    level_map = {"11": "Уровень 1", "21": "Уровень 2"}
    toc = {}
    body = document.xpath(".//w:body", namespaces=ns)[0]
    instrTexts = body.xpath(".//w:instrText", namespaces=ns)
    full_instr = "".join(it.text.strip() for it in instrTexts if it.text)
    if not full_instr.startswith('TOC \\o "1-3"'):
        print("TOC не найден")
        return toc
    sdtContent_ancestor = instrTexts[0].xpath("./ancestor::w:sdtContent[1]", namespaces=ns)[0]
    links = sdtContent_ancestor.xpath(".//w:hyperlink", namespaces=ns)
    for link in links:
        p = link.xpath("./ancestor::w:p[1]", namespaces=ns)[0]
        p_style = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=ns)
        style_val = p_style[0] if p_style else None

        text_parts = link.xpath(".//w:t/text()", namespaces=ns)
        full_text = "".join(text_parts).rstrip()

        m = re.search(r"([.\s]*)(\d+)$", full_text)
        if m:
            number = m.group(2)
            if 1 <= len(number) <= 3:
                full_text = full_text[:m.start()].rstrip()

        toc[full_text] = level_map.get(style_val, "Неизвестно")

    return toc


