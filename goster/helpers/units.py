class Units:
    TWIPS_PER_PT = 20
    TWIPS_PER_CM = 567
    EMU_PER_CM = 914400
    EMU_PER_PT = 12700
    HALF_POINTS_PER_PT = 2
    TWIPS_PER_LINE = 240

    @staticmethod
    def pt_to_twips(pt: float) -> int:
        return int(pt * Units.TWIPS_PER_PT)

    @staticmethod
    def twips_to_pt(twips: int) -> float:
        return twips / Units.TWIPS_PER_PT

    @staticmethod
    def cm_to_twips(cm: float) -> int:
        return int(cm * Units.TWIPS_PER_CM)

    @staticmethod
    def twips_to_cm(twips: int) -> float:
        return twips / Units.TWIPS_PER_CM

    @staticmethod
    def cm_to_emu(cm: float) -> int:
        return int(cm * Units.EMU_PER_CM)

    @staticmethod
    def emu_to_cm(emu: int) -> float:
        return emu / Units.EMU_PER_CM

    @staticmethod
    def pt_to_half_points(pt: float) -> int:
        return int(pt * Units.HALF_POINTS_PER_PT)

    @staticmethod
    def half_points_to_pt(half_points: int) -> float:
        return half_points / Units.HALF_POINTS_PER_PT

    @staticmethod
    def pt_to_emu(pt: float) -> int:
        return int(pt * Units.EMU_PER_PT)

    @staticmethod
    def line_spacing_to_twips(lines: float) -> int:
        return int(lines * Units.TWIPS_PER_LINE)

    @staticmethod
    def twips_to_line_spacing(twips: int) -> float:
        return twips / Units.TWIPS_PER_LINE
