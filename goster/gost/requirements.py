from ..helpers.units import Units


class GOSTRequirements:
    FONT_NAME = "Times New Roman"
    FONT_SIZE_PT = 14
    FONT_SIZE_MIN_PT = 12
    FONT_SIZE_MAX_PT = 14
    FONT_SIZE_HALF_POINTS = Units.pt_to_half_points(FONT_SIZE_PT)

    LINE_SPACING = 1.5
    LINE_SPACING_TWIPS = Units.line_spacing_to_twips(LINE_SPACING)

    FIRST_LINE_INDENT_CM = 1.25
    FIRST_LINE_INDENT_TWIPS = Units.cm_to_twips(FIRST_LINE_INDENT_CM)

    MARGIN_TOP_CM = 2.0
    MARGIN_BOTTOM_CM = 2.0
    MARGIN_LEFT_CM = 3.0
    MARGIN_RIGHT_MIN_CM = 1.0
    MARGIN_RIGHT_MAX_CM = 1.5

    MARGIN_TOP_TWIPS = Units.cm_to_twips(MARGIN_TOP_CM)
    MARGIN_BOTTOM_TWIPS = Units.cm_to_twips(MARGIN_BOTTOM_CM)
    MARGIN_LEFT_TWIPS = Units.cm_to_twips(MARGIN_LEFT_CM)
    MARGIN_RIGHT_MIN_TWIPS = Units.cm_to_twips(MARGIN_RIGHT_MIN_CM)
    MARGIN_RIGHT_MAX_TWIPS = Units.cm_to_twips(MARGIN_RIGHT_MAX_CM)

    ALIGNMENT_BODY = "both"
    ALIGNMENT_HEADING_STRUCTURAL = "center"
    ALIGNMENT_HEADING_CHAPTER = "center"
    ALIGNMENT_SUBHEADING = "left"

    FONT_COLOR = "000000"

    GOST_REFERENCE = "ГОСТ 7.32-2017"
