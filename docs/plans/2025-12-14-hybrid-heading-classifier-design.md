# Hybrid Heading Classifier Design

## Overview

Гибридная система классификации заголовков документа: эвристики + LLM для спорных случаев.

**Цель:** Улучшить точность определения заголовков без жёсткой привязки к стилям Word.

## Решения

- **Scope:** Только заголовки (тип, уровень 1-5, структурность)
- **Подход:** Гибридный — эвристики + LLM при confidence < 0.7
- **Fallback:** При недоступности LLM — использовать эвристики с пометкой `source="fallback"`
- **Инструменты:** LiteLLM (абстракция провайдеров) + BAML (structured output)

## Архитектура

```
┌─────────────────────────────────────────────────────────────┐
│                      DocumentParser                          │
├─────────────────────────────────────────────────────────────┤
│  1. Извлечь сырые параграфы из docx                         │
│  2. Для каждого параграфа:                                  │
│     ├─ HeuristicClassifier.classify() → (result, confidence)│
│     ├─ if confidence < 0.7:                                 │
│     │    └─ LLMClassifier.classify() → result               │
│     └─ else: использовать результат эвристик               │
│  3. Построить DocumentModel из классифицированных элементов │
└─────────────────────────────────────────────────────────────┘
```

## Структура файлов

```
goster/
├── core/
│   ├── parser.py          # Адаптированный парсер
│   ├── classifier.py      # ClassificationResult, HeuristicClassifier, HybridClassifier
│   └── llm_classifier.py  # LLMClassifier
└── baml_src/
    └── heading.baml       # BAML схема для классификации
```

## Типы данных

```python
@dataclass
class ClassificationContext:
    text: str
    style_name: str | None
    prev_texts: list[str]  # предыдущие 2-3 параграфа

@dataclass
class ClassificationResult:
    element_type: Literal["heading", "paragraph"]
    confidence: float  # 0.0 - 1.0
    heading_level: int | None  # 1-5 для заголовков
    is_structural: bool  # введение, заключение, содержание...
    source: Literal["heuristic", "llm", "fallback"]

class HeadingClassifier(Protocol):
    def classify(self, ctx: ClassificationContext) -> ClassificationResult
```

## HeuristicClassifier

Перенос логики из существующих методов `_is_heading_like()` и `_get_heading_level()` с добавлением confidence score:

| Паттерн | Confidence | Пример |
|---------|------------|--------|
| Структурные ключевые слова | 0.95 | ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ |
| ГЛАВА N | 0.95 | ГЛАВА 1 |
| Нумерованные N.N Текст | 0.85 | 1.2 Актуальность |
| Стиль Heading N | 0.75 | style="Heading 2" |
| Короткий с заглавной | 0.40 | Анализ результатов |
| Остальное | 0.95 (paragraph) | Обычный текст... |

## BAML Schema

```baml
class HeadingClassification {
  is_heading bool
  level int?
  is_structural bool
  reasoning string
}

function ClassifyHeading(text: string, context: string[]) -> HeadingClassification {
  client LiteLLM
  prompt #"
    Ты эксперт по оформлению ВКР по ГОСТ 7.32-2017.
    Определи, является ли текст заголовком раздела.

    Контекст (предыдущие параграфы):
    {% for ctx in context %}
    - {{ ctx }}
    {% endfor %}

    Текст для классификации:
    {{ text }}

    Правила:
    - Уровень 1: главы, введение, заключение, содержание
    - Уровень 2: подразделы (1.1, 1.2)
    - Уровень 3+: пункты (1.1.1, 1.1.2)
    - Структурные: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СОДЕРЖАНИЕ, ПРИЛОЖЕНИЕ

    {{ ctx.output_format }}
  "#
}
```

## Интеграция в parser.py

```python
class DocumentParser:
    def __init__(self, path: Path | str, classifier: HeadingClassifier | None = None):
        ...
        self._classifier = classifier or HybridHeadingClassifier()

    def _parse_paragraph(self, para: Paragraph, index: int) -> Element:
        context = ClassificationContext(
            text=para.text.strip(),
            style_name=para.style.name if para.style else None,
            prev_texts=self._get_prev_texts(index, count=3),
        )

        result = self._classifier.classify(context)

        if result.element_type == "heading":
            heading = HeadingElement(index=index, raw=para, level=result.heading_level)
            heading.is_structural = result.is_structural
            heading.classification_source = result.source
            ...
```

## Конфигурация

```python
# Только эвристики (тесты, оффлайн)
parser = DocumentParser(path, classifier=HeuristicClassifier())

# Гибридный режим (production)
parser = DocumentParser(path, classifier=HybridHeadingClassifier(
    confidence_threshold=0.7,
    llm_classifier=LLMClassifier()
))
```

## Зависимости

```toml
[project.dependencies]
litellm = ">=1.0.0"
baml = ">=0.1.0"
```
