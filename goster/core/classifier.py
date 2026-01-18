from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Literal, Protocol

HEADING_NUMBERED_PATTERN = re.compile(r"^\d+(\.\d+)+\s+[А-ЯA-Z]")
HEADING_SINGLE_NUMBER_PATTERN = re.compile(r"^\d+\s+[А-ЯA-Z]")
CHAPTER_PATTERN = re.compile(r"^ГЛАВА\s+(\d+)", re.IGNORECASE)
NUMBER_PREFIX_PATTERN = re.compile(r"^(\d+(?:\.\d+)*)")

STRUCTURAL_KEYWORDS = frozenset([
    "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СОДЕРЖАНИЕ", "РЕФЕРАТ", "АННОТАЦИЯ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ЛИТЕРАТУРЫ",
    "БИБЛИОГРАФИЧЕСКИЙ СПИСОК", "ОГЛАВЛЕНИЕ", "ПРИЛОЖЕНИЕ",
    "СПИСОК СОКРАЩЕНИЙ", "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ", "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ",
    "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ", "КЛЮЧЕВЫЕ СЛОВА",
])

STRUCTURAL_LEVEL_1 = frozenset([
    "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ",
    "РЕФЕРАТ", "АННОТАЦИЯ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "СПИСОК ЛИТЕРАТУРЫ", "БИБЛИОГРАФИЧЕСКИЙ СПИСОК",
])


@dataclass
class ClassificationContext:
    text: str
    style_name: str | None = None
    prev_texts: list[str] = field(default_factory=list)


@dataclass
class ClassificationResult:
    element_type: Literal["heading", "paragraph"]
    confidence: float
    heading_level: int | None = None
    is_structural: bool = False
    source: Literal["heuristic", "llm", "fallback"] = "heuristic"


class HeadingClassifier(Protocol):
    def classify(self, ctx: ClassificationContext) -> ClassificationResult: ...


class HeuristicClassifier:
    def classify(self, ctx: ClassificationContext) -> ClassificationResult:
        text = ctx.text.strip()
        text_upper = text.upper()

        if not text:
            return ClassificationResult(
                element_type="paragraph",
                confidence=1.0,
                source="heuristic"
            )

        for kw in STRUCTURAL_KEYWORDS:
            if text_upper.startswith(kw) or text_upper == kw:
                is_level_1 = any(text_upper.startswith(k) for k in STRUCTURAL_LEVEL_1)
                return ClassificationResult(
                    element_type="heading",
                    confidence=0.95,
                    heading_level=1 if is_level_1 else 2,
                    is_structural=True,
                    source="heuristic"
                )

        chapter_match = CHAPTER_PATTERN.match(text_upper)
        if chapter_match:
            return ClassificationResult(
                element_type="heading",
                confidence=0.95,
                heading_level=1,
                is_structural=False,
                source="heuristic"
            )

        if HEADING_NUMBERED_PATTERN.match(text):
            num_match = NUMBER_PREFIX_PATTERN.match(text)
            level = len(num_match.group(1).split(".")) if num_match else 2
            return ClassificationResult(
                element_type="heading",
                confidence=0.85,
                heading_level=level,
                is_structural=False,
                source="heuristic"
            )

        if HEADING_SINGLE_NUMBER_PATTERN.match(text):
            return ClassificationResult(
                element_type="heading",
                confidence=0.80,
                heading_level=1,
                is_structural=False,
                source="heuristic"
            )

        if ctx.style_name and ctx.style_name.startswith("Heading"):
            try:
                level = int(ctx.style_name.split()[-1])
                confidence = 0.75 if len(text) <= 100 else 0.50
                return ClassificationResult(
                    element_type="heading",
                    confidence=confidence,
                    heading_level=level,
                    is_structural=False,
                    source="heuristic"
                )
            except (ValueError, IndexError):
                pass

        if len(text) > 150:
            return ClassificationResult(
                element_type="paragraph",
                confidence=0.90,
                source="heuristic"
            )

        if " — " in text or " - " in text:
            return ClassificationResult(
                element_type="paragraph",
                confidence=0.85,
                source="heuristic"
            )

        if text_upper.startswith("ПРОДОЛЖЕНИЕ ТАБЛИЦЫ"):
            return ClassificationResult(
                element_type="paragraph",
                confidence=0.95,
                source="heuristic"
            )

        if text.endswith(":"):
            return ClassificationResult(
                element_type="paragraph",
                confidence=0.80,
                source="heuristic"
            )

        if len(text) < 80 and text and text[0].isupper() and not text.endswith("."):
            return ClassificationResult(
                element_type="heading",
                confidence=0.40,
                heading_level=2,
                is_structural=False,
                source="heuristic"
            )

        return ClassificationResult(
            element_type="paragraph",
            confidence=0.90,
            source="heuristic"
        )
