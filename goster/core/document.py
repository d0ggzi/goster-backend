from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from docx import Document as DocumentFactory
from docx.document import Document as DocumentObject
from docx.text.paragraph import Paragraph

from .model import DocumentModel, ParagraphElement, HeadingElement
from .parser import DocumentParser


@dataclass
class DocumentContext:
    input_path: Path
    use_llm: bool = False
    document: DocumentObject = field(init=False)
    model: DocumentModel = field(init=False)
    _element_to_paragraph: dict[int, Paragraph] = field(init=False, default_factory=dict)

    def __post_init__(self):
        self.input_path = Path(self.input_path)
        self.document = DocumentFactory(self.input_path)
        self._parse_model()

    def _parse_model(self):
        classifier = None
        if self.use_llm:
            from .llm_classifier import HybridHeadingClassifier
            classifier = HybridHeadingClassifier(confidence_threshold=0.5)
        parser = DocumentParser.from_document(self.input_path, self.document, classifier=classifier)
        self.model = parser.parse()
        self._build_element_mapping()

    def _build_element_mapping(self):
        for element in self.model.elements:
            if isinstance(element, (ParagraphElement, HeadingElement)):
                self._element_to_paragraph[element.index] = element.raw

    def get_paragraph_at(self, element_index: int) -> Paragraph | None:
        return self._element_to_paragraph.get(element_index)

    def save(self, output_path: Path):
        self.document.save(output_path)
