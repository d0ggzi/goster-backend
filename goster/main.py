from pathlib import Path

from dotenv import load_dotenv

from .core import DocumentContext, ValidationPipeline
from .rules import (
    FontRule,
    AlignmentRule,
    IndentRule,
    SpacingRule,
    ParagraphSpacingRule,
    TextColorRule,
    NoUnderlineRule,
    ListFormatRule,
    HeadingNumberingRule,
    HeadingCaseRule,
    HeadingAlignmentRule,
    HeadingPunctuationRule,
    HeadingBoldRule,
    HeadingNewPageRule,
    SubheadingCaseRule,
    RequiredSectionsRule,
    TableReferenceRule,
    FigureReferenceRule,
    TableCaptionRule,
    FigureCaptionRule,
    FormulaReferenceRule,
    CitationFormatRule,
    MarginsRule,
    PageNumberingRule,
    AppendixRule,
)


def create_default_pipeline() -> ValidationPipeline:
    pipeline = ValidationPipeline()
    pipeline.add_rules(
        MarginsRule(),
        FontRule(),
        TextColorRule(),
        AlignmentRule(),
        IndentRule(),
        SpacingRule(),
        ParagraphSpacingRule(),
        NoUnderlineRule(),
        ListFormatRule(),
        HeadingNumberingRule(),
        HeadingCaseRule(),
        HeadingAlignmentRule(),
        HeadingPunctuationRule(),
        HeadingBoldRule(),
        HeadingNewPageRule(),
        SubheadingCaseRule(),
        RequiredSectionsRule(),
        TableReferenceRule(),
        FigureReferenceRule(),
        TableCaptionRule(),
        FigureCaptionRule(),
        FormulaReferenceRule(),
        CitationFormatRule(),
        AppendixRule(),
        PageNumberingRule(),
    )
    return pipeline


def validate_document(input_path: str | Path, use_llm: bool = False) -> str:
    if use_llm:
        load_dotenv()
    ctx = DocumentContext(input_path, use_llm=use_llm)
    pipeline = create_default_pipeline()
    report = pipeline.validate(ctx)
    return report.detailed_report()


def validate_and_fix_document(
    input_path: str | Path,
    output_path: str | Path,
    use_llm: bool = False,
) -> str:
    if use_llm:
        load_dotenv()
    ctx = DocumentContext(input_path, use_llm=use_llm)
    pipeline = create_default_pipeline()
    report = pipeline.validate_and_highlight(ctx)
    ctx.save(Path(output_path))
    return report.detailed_report()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="ГОСТ 7.32-2017 валидатор документов")
    parser.add_argument("input", help="Входной .docx файл")
    parser.add_argument("-o", "--output", help="Выходной файл с подсветкой ошибок")
    parser.add_argument("--llm", action="store_true", help="Использовать LLM для классификации заголовков")

    args = parser.parse_args()

    if args.output:
        report = validate_and_fix_document(args.input, args.output, use_llm=args.llm)
        print(f"Документ сохранён: {args.output}")
    else:
        report = validate_document(args.input, use_llm=args.llm)

    print(report)
