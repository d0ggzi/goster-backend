from __future__ import annotations

import logging

from .classifier import ClassificationContext, ClassificationResult

logger = logging.getLogger(__name__)


class LLMClassifier:
    def __init__(self):
        self._client = None

    def _get_client(self):
        if self._client is None:
            from ..baml_client.baml_client.sync_client import b

            self._client = b
        return self._client

    def classify(self, ctx: ClassificationContext) -> ClassificationResult | None:
        try:
            client = self._get_client()
            result = client.ClassifyHeading(
                text=ctx.text, prev_paragraphs=ctx.prev_texts or []
            )

            logger.debug(
                "LLM classification: is_heading=%s, level=%s, reasoning=%s",
                result.is_heading,
                result.level,
                result.reasoning,
            )

            return ClassificationResult(
                element_type="heading" if result.is_heading else "paragraph",
                confidence=0.90,
                heading_level=result.level,
                is_structural=result.is_structural,
                source="llm",
            )

        except Exception as e:
            logger.warning("LLM classification failed: %s", e)
            return None


class HybridHeadingClassifier:
    def __init__(
        self,
        confidence_threshold: float = 0.7,
        llm_classifier: LLMClassifier | None = None,
        use_llm: bool = True,
    ):
        from .classifier import HeuristicClassifier

        self._heuristic = HeuristicClassifier()
        self._llm = llm_classifier or LLMClassifier() if use_llm else None
        self._threshold = confidence_threshold

    def classify(self, ctx: ClassificationContext) -> ClassificationResult:
        result = self._heuristic.classify(ctx)

        if result.confidence >= self._threshold:
            return result

        if self._llm is not None:
            llm_result = self._llm.classify(ctx)
            if llm_result is not None:
                return llm_result

        return ClassificationResult(
            element_type=result.element_type,
            confidence=result.confidence,
            heading_level=result.heading_level,
            is_structural=result.is_structural,
            source="fallback",
        )
