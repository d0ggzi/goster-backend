// Heading classification schema for GOST 7.32-2017 documents

class HeadingClassification {
  is_heading bool @description("True if text is a section heading")
  level int? @description("Heading level 1-5, null if not a heading")
  is_structural bool @description("Structural section: introduction, conclusion, contents, appendix, references")
  reasoning string @description("Brief explanation of the decision")
}

function ClassifyHeading(text: string, prev_paragraphs: string[]) -> HeadingClassification {
  client LiteLLM

  prompt #"
    You are an expert in Russian academic document formatting (GOST 7.32-2017).

    Determine if the given text is a section heading.

    Context (previous paragraphs):
    {{ prev_paragraphs }}

    Text to classify:
    {{ text }}

    Classification rules:
    - Level 1: chapters (ГЛАВА N), structural sections (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СОДЕРЖАНИЕ)
    - Level 2: subsections (1.1, 1.2, 2.1)
    - Level 3+: sub-subsections (1.1.1, 1.1.2, 2.1.1)
    - Structural sections: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СОДЕРЖАНИЕ, РЕФЕРАТ, АННОТАЦИЯ,
      СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ, СПИСОК ЛИТЕРАТУРЫ, ПРИЛОЖЕНИЕ

    Signs of a heading:
    - Short text (usually < 100 chars)
    - Starts with number (1, 1.1, 1.1.1) or "ГЛАВА"
    - Contains structural keywords in uppercase
    - Does not end with period
    - Often in uppercase or title case

    Signs of regular paragraph:
    - Long text with multiple sentences
    - Ends with period
    - Contains detailed explanations

    {{ ctx.output_format }}
  "#
}
