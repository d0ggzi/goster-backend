from lxml import etree

HEAD1_TITLES = [
    "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ",
    "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ",
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "ПРИЛОЖЕНИЯ",
]


def get_next_key(d: dict, key: str):
    keys = list(d.keys())
    if key not in keys:
        return None
    idx = keys.index(key)
    return keys[idx + 1] if idx + 1 < len(keys) else None


def fix_sections(toc, document, styles):
    keys = list(toc.keys())
    for section_title in keys:
        title_upper = section_title.upper()
        next_title = get_next_key(toc, section_title)
        if "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ" in title_upper:
            fix_list_of_abbreviations_and_symbols(document, styles, section_title, next_title)
        elif "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ" in title_upper:
            fix_terms_and_definitions(document, styles, section_title, next_title)


def fix_list_of_abbreviations_and_symbols(document, styles, section_title, next_title):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    sdt = document.xpath(".//w:body/w:sdt", namespaces=ns)[0]
    title_paragraph = sdt.xpath(f'following-sibling::w:p[w:r/w:t="{section_title}"]', namespaces=ns)[0]
    paragraph = title_paragraph.getnext()
    while paragraph is not None:
        paragraph_text = "".join(paragraph.xpath(".//w:t/text()", namespaces=ns))
        if paragraph_text == next_title:
            break
        pPr_list = paragraph.xpath(".//w:pPr", namespaces=ns)
        for pPr in pPr_list:
            paragraph.remove(pPr)
        pPr = etree.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
        pStyle = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
        pStyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "NormalNoIndent")
        paragraph.insert(0, pPr)
        print(paragraph_text)
        paragraph = paragraph.getnext()


def fix_terms_and_definitions(document, styles, section_title, next_title):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    sdt = document.xpath(".//w:body/w:sdt", namespaces=ns)[0]
    title_paragraph = sdt.xpath(f'following-sibling::w:p[w:r/w:t="{section_title}"]', namespaces=ns)[0]
    paragraph = title_paragraph.getnext()
    while paragraph is not None:
        paragraph_text = "".join(paragraph.xpath(".//w:t/text()", namespaces=ns))
        if paragraph_text == next_title:
            break
        pPr_list = paragraph.xpath(".//w:pPr", namespaces=ns)
        for pPr in pPr_list:
            paragraph.remove(pPr)
        pPr = etree.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
        pStyle = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
        pStyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "NormalNoIndent")
        paragraph.insert(0, pPr)
        print(paragraph_text)
        paragraph = paragraph.getnext()


def fix_struct_headings(document, toc):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    sdt = document.xpath(".//w:body/w:sdt", namespaces=ns)[0]
    for head in HEAD1_TITLES:
        if head in toc.keys():
            paragraph = sdt.xpath(f'following-sibling::w:p[w:r/w:t="{head}"]', namespaces=ns)
            if paragraph:
                paragraph = paragraph[0]
                pPr_list = paragraph.xpath(".//w:pPr", namespaces=ns)
                for pPr in pPr_list:
                    paragraph.remove(pPr)
                pPr = etree.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
                pStyle = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
                pStyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "HeadingStruct")
                paragraph.insert(0, pPr)


def fix_level1_headings(document, toc):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    sdt = document.xpath(".//w:body/w:sdt", namespaces=ns)[0]
    for head in toc:
        paragraph = sdt.xpath(f'following-sibling::w:p[w:r/w:t="{head}"]', namespaces=ns)
        if paragraph:
            paragraph = paragraph[0]
            pPr_list = paragraph.xpath(".//w:pPr", namespaces=ns)
            for pPr in pPr_list:
                paragraph.remove(pPr)
            pPr = etree.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
            pStyle = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
            pStyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "Head1")
            paragraph.insert(0, pPr)


def fix_main_sections(main_headings, document, style="Normal"):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    sdt = document.xpath(".//w:body/w:sdt", namespaces=ns)[0]
    for i, head in enumerate(main_headings):
        paragraph_list = sdt.xpath(f'following-sibling::w:p[w:r/w:t="{head}"]', namespaces=ns)
        if not paragraph_list:
            continue
        paragraph = paragraph_list[0]
        next_head = main_headings[i + 1] if i + 1 < len(main_headings) else None
        current = paragraph.getnext()
        while current is not None:
            paragraph_text = "".join(current.xpath(".//w:t/text()", namespaces=ns))
            if paragraph_text == next_head:
                break
            for pPr in current.xpath(".//w:pPr", namespaces=ns):
                current.remove(pPr)
            pPr = etree.SubElement(current, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
            pStyle = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
            pStyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", style)
            current = current.getnext()


def get_head1_toc(toc):
    result = []
    for title, level in toc.items():
        if level == "Уровень 1" and title.upper() not in [h.upper() for h in HEAD1_TITLES]:
            result.append(title)
    return result


def get_head2_toc(toc):
    result = []
    for title, level in toc.items():
        if level == "Уровень 2":
            result.append(title)
    return result


def get_main_headings(toc):
    result = []
    for title in toc.keys():
        if title not in HEAD1_TITLES:
            result.append(title)
    return result
