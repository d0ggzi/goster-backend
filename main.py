import os
import zipfile

from lxml import etree

from docx_handler import save_docx, get_table_of_contents, get_document
from style_utils import update_normal_style, update_heading_1_style, create_normal_no_indent_style, get_styles, \
    create_heading_struct_style, get_numbering, create_head1_style
from section_formatter import fix_sections, fix_struct_headings, fix_level1_headings, \
    get_head1_toc, get_head2_toc, get_main_headings, fix_main_sections


def xml_to_string(xml_element):
    xml_string =  etree.tostring(
        xml_element,
        encoding="unicode",
        pretty_print=True
    )
    print(xml_string)

def save_xml(xml_element, output_path):
    xml_bytes = etree.tostring(
        xml_element,
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=True
    )

    with open(output_path, "wb") as f:
        f.write(xml_bytes)

    print(f"XML успешно сохранён в {output_path}")

def extract_docx_xml(docx_path, output_dir="docx_xml"):
    os.makedirs(output_dir, exist_ok=True)
    xml_files = {}

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        for name in zip_ref.namelist():
            if name.endswith(".xml"):
                xml_bytes = zip_ref.read(name)


                xml_element = etree.fromstring(xml_bytes)
                xml_files[name] = xml_element


                out_path = os.path.join(output_dir, name.replace("/", "_"))
                with open(out_path, "wb") as f:
                    f.write(etree.tostring(xml_element, pretty_print=True, encoding="utf-8"))

                print(f"Сохранён: {out_path}")

    return xml_files

def main():
    input_path = "push.docx"
    output_path = "qwerty.docx"
    document = get_document(input_path)
    toc = get_table_of_contents(document)
    print(toc)
    head1_toc = get_head1_toc(toc)
    print(head1_toc)
    head2_toc = get_head2_toc(toc)
    print(head2_toc)
    main_headings = get_main_headings(toc)
    print(main_headings)
    styles = get_styles(input_path)
    update_normal_style(styles)
    create_normal_no_indent_style(styles)
    create_heading_struct_style(styles)
    create_head1_style(styles)
    fix_sections(toc, document, styles)
    fix_main_sections(main_headings, document)
    fix_struct_headings(document, toc)
    fix_level1_headings(document, head1_toc)

    save_docx(input_path, output_path, styles, document, None)


if __name__ == "__main__":
    main()
