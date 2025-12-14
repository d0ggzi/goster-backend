from zipfile import ZipFile
from lxml import etree


def get_styles(input_file):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(input_file) as docx_zip:
        xml_content = docx_zip.read("word/styles.xml")
    root = etree.fromstring(xml_content)
    return root


def load_default_numbering():
    with open("../../../assets/default-numbering.xml", "rb") as f:
        return etree.fromstring(f.read())


def get_numbering(input_file):
    numbering_path = "word/numbering.xml"
    with ZipFile(input_file) as docx_zip:
        if numbering_path in docx_zip.namelist():
            return etree.fromstring(docx_zip.read(numbering_path))
        print("⚠ numbering.xml отсутствует, используем стандартный шаблон.")
        return load_default_numbering()


def update_heading_1_style(root):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    styles = root.xpath(".//w:style[w:name/@w:val='heading 1']", namespaces=ns)
    if not styles:
        print("❌ Heading 1 не найден!")
        return
    style = styles[0]
    print("✅ Heading 1 найден, styleId:",
          style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId"))
    pPr = style.find("w:pPr", namespaces=ns)
    if pPr is None:
        pPr = etree.SubElement(style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    jc = pPr.find("w:jc", namespaces=ns)
    if jc is None:
        jc = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
    jc.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "both")
    ind = pPr.find("w:ind", namespaces=ns)
    if ind is None:
        ind = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
    ind.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine", "0")
    numPr = pPr.find("w:numPr", namespaces=ns)
    if numPr is None:
        numPr = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
    ilvl = numPr.find("w:ilvl", namespaces=ns)
    if ilvl is None:
        ilvl = etree.SubElement(numPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl")
    ilvl.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "0")
    numId = numPr.find("w:numId", namespaces=ns)
    if numId is None:
        numId = etree.SubElement(numPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId")
    numId.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "1")
    rPr = style.find("w:rPr", namespaces=ns)
    if rPr is None:
        rPr = etree.SubElement(style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    old_rFonts = rPr.find("w:rFonts", namespaces=ns)
    if old_rFonts is not None:
        rPr.remove(old_rFonts)
    rFonts = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii", "Times New Roman")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi", "Times New Roman")
    sz = rPr.find("w:sz", namespaces=ns)
    if sz is None:
        sz = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz")
    sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "28")
    b = rPr.find("w:b", namespaces=ns)
    if b is None:
        b = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b")
    b.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "true")
    caps = rPr.find("w:caps", namespaces=ns)
    if caps is None:
        caps = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}caps")
    caps.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "true")
    color = rPr.find("w:color", namespaces=ns)
    if color is None:
        color = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
    else:
        for attr in list(color.attrib.keys()):
            if "theme" in attr.lower():
                del color.attrib[attr]
    color.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "000000")
    print("✅ Heading 1 приведён к ГОСТу и использует нумерацию 9999")
    return root


def update_normal_style(root):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    styles = root.xpath(".//w:style[w:name/@w:val='Normal']", namespaces=ns)
    if not styles:
        print("❌ Normal не найден!")
        return
    style = styles[0]
    print("✅ Normal найден, styleId:",
          style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId"))
    pPr = style.find("w:pPr", namespaces=ns)
    if pPr is None:
        pPr = etree.SubElement(style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    jc = pPr.find("w:jc", namespaces=ns)
    if jc is None:
        jc = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
    jc.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "both")
    ind = pPr.find("w:ind", namespaces=ns)
    if ind is None:
        ind = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
    ind.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine", "708")
    spacing = pPr.find("w:spacing", namespaces=ns)
    if spacing is None:
        spacing = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing")
    spacing.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line", "360")
    spacing.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule", "auto")
    rPr = style.find("w:rPr", namespaces=ns)
    if rPr is None:
        rPr = etree.SubElement(style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    rFonts = rPr.find("w:rFonts", namespaces=ns)
    if rFonts is not None:
        rPr.remove(rFonts)
    rFonts = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii", "Times New Roman")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi", "Times New Roman")
    sz = rPr.find("w:sz", namespaces=ns)
    if sz is None:
        sz = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz")
    sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "28")
    color = rPr.find("w:color", namespaces=ns)
    if color is None:
        color = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
    else:
        for attr in list(color.attrib.keys()):
            if "theme" in attr.lower():
                del color.attrib[attr]
    color.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "000000")
    print("✅ Стиль Normal приведён к ГОСТу")
    return root


def create_normal_no_indent_style(root):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    normal_styles = root.xpath(".//w:style[w:name/@w:val='Normal']", namespaces=ns)
    if not normal_styles:
        print("❌ Стиль Normal не найден!")
        return
    normal = normal_styles[0]
    new_style = etree.SubElement(root, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style")
    new_style.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "paragraph")
    new_style.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", "NormalNoIndent")
    name = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")
    name.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "NormalNoIndent")
    basedOn = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}basedOn")
    basedOn.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "Normal")
    pPr = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    jc = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
    jc.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "both")
    ind = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
    ind.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine", "0")
    spacing = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing")
    spacing.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line", "360")
    spacing.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule", "auto")
    rPr = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    rFonts = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii", "Times New Roman")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi", "Times New Roman")
    sz = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz")
    sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "28")
    color = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
    color.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "000000")
    print("✅ Стиль NormalNoIndent создан с ГОСТ-настройками:")
    print("   Шрифт: Times New Roman, 14 пт, 1.5 строки, выравнивание по ширине, без отступа.")
    return root


def create_heading_struct_style(root):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    new_style = etree.SubElement(root, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style")
    new_style.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "paragraph")
    new_style.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", "HeadingStruct")
    name = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")
    name.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "HeadingStruct")
    pPr = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    jc = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
    jc.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "center")
    ind = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
    ind.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine", "0")
    rPr = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    rFonts = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii", "Times New Roman")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi", "Times New Roman")
    sz = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz")
    sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "28")
    b = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b")
    b.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "true")
    caps = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}caps")
    caps.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "true")
    color = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
    color.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "000000")
    print("✅ Стиль HeadingStruct создан с заданными параметрами.")
    return root


def create_head1_style(root):
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    new_style = etree.SubElement(root, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style")
    new_style.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "paragraph")
    new_style.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", "Head1")
    name = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")
    name.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "Head1")
    pPr = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    jc = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
    jc.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "both")
    ind = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
    ind.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine", "708")
    numPr = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
    ilvl = etree.SubElement(numPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl")
    ilvl.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "0")
    numId = etree.SubElement(numPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId")
    numId.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "1")
    rPr = etree.SubElement(new_style, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    rFonts = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii", "Times New Roman")
    rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi", "Times New Roman")
    sz = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz")
    sz.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "28")
    b = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b")
    b.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "true")
    caps = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}caps")
    caps.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "true")
    color = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
    color.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "000000")
    print("✅ Стиль Head1 создан с заданными параметрами и абзацным отступом 1,25 см.")
    return root
