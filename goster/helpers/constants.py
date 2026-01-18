from docx.enum.text import WD_ALIGN_PARAGRAPH


ALIGNMENT_MAP = {
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}

ALIGNMENT_NAMES = {
    WD_ALIGN_PARAGRAPH.JUSTIFY: "по ширине",
    WD_ALIGN_PARAGRAPH.LEFT: "по левому краю",
    WD_ALIGN_PARAGRAPH.CENTER: "по центру",
    WD_ALIGN_PARAGRAPH.RIGHT: "по правому краю",
    None: "не задано",
}
