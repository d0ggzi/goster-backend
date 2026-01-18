from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile
import shutil

from lxml import etree


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def save_docx_with_xml(
    input_path: Path,
    output_path: Path,
    document_xml: etree._Element | None = None,
    styles_xml: etree._Element | None = None,
    numbering_xml: etree._Element | None = None,
):
    document_bytes = (
        etree.tostring(
            document_xml, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )
        if document_xml is not None
        else None
    )
    styles_bytes = (
        etree.tostring(
            styles_xml, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )
        if styles_xml is not None
        else None
    )
    numbering_bytes = (
        etree.tostring(
            numbering_xml, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )
        if numbering_xml is not None
        else None
    )

    temp_path = Path(str(output_path) + "_temp.zip")

    with ZipFile(input_path, "r") as zin, ZipFile(temp_path, "w") as zout:
        has_numbering = False

        for item in zin.infolist():
            if item.filename == "word/document.xml" and document_bytes:
                zout.writestr("word/document.xml", document_bytes)
            elif item.filename == "word/styles.xml" and styles_bytes:
                zout.writestr("word/styles.xml", styles_bytes)
            elif item.filename == "word/numbering.xml":
                has_numbering = True
                if numbering_bytes:
                    zout.writestr("word/numbering.xml", numbering_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))
            else:
                zout.writestr(item, zin.read(item.filename))

        if not has_numbering and numbering_bytes:
            zout.writestr("word/numbering.xml", numbering_bytes)

    shutil.move(temp_path, output_path)


def get_or_create_element(parent: etree._Element, tag: str) -> etree._Element:
    full_tag = f"{W}{tag}"
    element = parent.find(f"w:{tag}", namespaces=NS)
    if element is None:
        element = etree.SubElement(parent, full_tag)
    return element


def set_element_value(parent: etree._Element, tag: str, value: str) -> etree._Element:
    element = get_or_create_element(parent, tag)
    element.set(f"{W}val", value)
    return element
