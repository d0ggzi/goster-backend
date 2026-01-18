from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Iterator

from docx.text.paragraph import Paragraph
from docx.table import Table


STRUCTURAL_KEYWORDS = frozenset([
    "РЕФЕРАТ",
    "АННОТАЦИЯ",
    "КЛЮЧЕВЫЕ СЛОВА",
    "СОДЕРЖАНИЕ",
    "ОГЛАВЛЕНИЕ",
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "СПИСОК ЛИТЕРАТУРЫ",
    "БИБЛИОГРАФИЧЕСКИЙ СПИСОК",
    "СПИСОК СОКРАЩЕНИЙ",
    "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ",
    "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ",
    "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ",
    "ПРИЛОЖЕНИЕ",
])


class ElementType(Enum):
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    TABLE = "table"
    FIGURE = "figure"
    FORMULA = "formula"


@dataclass
class Element:
    type: ElementType
    index: int
    raw: Any

    @property
    def text(self) -> str:
        if hasattr(self.raw, "text"):
            return self.raw.text or ""
        return ""


@dataclass
class ParagraphElement(Element):
    type: ElementType = field(default=ElementType.PARAGRAPH, init=False)

    @property
    def paragraph(self) -> Paragraph:
        return self.raw


@dataclass
class HeadingElement(Element):
    type: ElementType = field(default=ElementType.HEADING, init=False)
    level: int = 0
    number: str | None = None
    classification_source: str = "heuristic"
    classification_confidence: float = 1.0

    @property
    def paragraph(self) -> Paragraph:
        return self.raw

    @property
    def is_structural(self) -> bool:
        text_upper = self.text.strip().upper()
        return any(text_upper.startswith(kw) for kw in STRUCTURAL_KEYWORDS)


class ReferencedElementMixin:
    referenced_at: list[int]
    index: int

    @property
    def has_reference_before(self) -> bool:
        if not self.referenced_at:
            return False
        return any(ref_idx < self.index for ref_idx in self.referenced_at)


@dataclass
class TableElement(ReferencedElementMixin, Element):
    type: ElementType = field(default=ElementType.TABLE, init=False)
    caption: str | None = None
    number: str | None = None
    caption_paragraph_index: int | None = None
    referenced_at: list[int] = field(default_factory=list)

    @property
    def table(self) -> Table:
        return self.raw


@dataclass
class FigureElement(ReferencedElementMixin, Element):
    type: ElementType = field(default=ElementType.FIGURE, init=False)
    caption: str | None = None
    number: str | None = None
    referenced_at: list[int] = field(default_factory=list)


@dataclass
class FormulaElement(Element):
    type: ElementType = field(default=ElementType.FORMULA, init=False)
    number: str | None = None


@dataclass
class Section:
    heading: HeadingElement
    level: int
    parent: Section | None = None
    children: list[Section] = field(default_factory=list)
    elements: list[Element] = field(default_factory=list)

    @property
    def number(self) -> str | None:
        return self.heading.number

    @property
    def title(self) -> str:
        return self.heading.text

    def iter_elements(self) -> Iterator[Element]:
        yield self.heading
        yield from self.elements
        for child in self.children:
            yield from child.iter_elements()

    def __str__(self) -> str:
        return f"Секция: level={self.level}, title={self.heading.text[:50]!r}"

    def __repr__(self) -> str:
        return f"Section(level={self.level}, title={self.heading.text[:30]!r}, children={len(self.children)})"


@dataclass
class DocumentModel:
    sections: list[Section] = field(default_factory=list)
    elements: list[Element] = field(default_factory=list)
    tables: list[TableElement] = field(default_factory=list)
    figures: list[FigureElement] = field(default_factory=list)
    formulas: list[FormulaElement] = field(default_factory=list)
    headings: list[HeadingElement] = field(default_factory=list)
    paragraphs: list[ParagraphElement] = field(default_factory=list)
    preamble_elements: list[Element] = field(default_factory=list)

    _element_to_section: dict[int, Section] = field(default_factory=dict, repr=False)

    def get_section_for_element(self, element: Element) -> Section | None:
        return self._element_to_section.get(element.index)

    def get_element_at(self, index: int) -> Element | None:
        if 0 <= index < len(self.elements):
            return self.elements[index]
        return None

    def iter_paragraphs(self) -> Iterator[ParagraphElement]:
        yield from self.paragraphs

    def iter_headings(self) -> Iterator[HeadingElement]:
        yield from self.headings

    def iter_tables(self) -> Iterator[TableElement]:
        yield from self.tables

    def iter_figures(self) -> Iterator[FigureElement]:
        yield from self.figures

    def iter_all(self) -> Iterator[Element]:
        yield from self.elements

    def get_heading_hierarchy(self) -> list[tuple[int, HeadingElement]]:
        return [(h.level, h) for h in self.headings]
