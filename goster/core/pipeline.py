from .document import DocumentContext
from .highlighter import DocumentHighlighter
from .report import ValidationReport
from .rule import BaseRule


class ValidationPipeline:
    def __init__(self):
        self._rules: list[BaseRule] = []

    def add_rule(self, rule: BaseRule) -> "ValidationPipeline":
        self._rules.append(rule)
        return self

    def add_rules(self, *rules: BaseRule) -> "ValidationPipeline":
        self._rules.extend(rules)
        return self

    @property
    def rules(self) -> list[BaseRule]:
        return self._rules.copy()

    def validate(self, ctx: DocumentContext) -> ValidationReport:
        report = ValidationReport()

        for rule in self._rules:
            errors = rule.validate(ctx)
            report.add_errors(errors)

        return report

    def validate_and_fix(self, ctx: DocumentContext) -> ValidationReport:
        report = ValidationReport()

        for rule in self._rules:
            errors, fixes = rule.validate_and_fix(ctx)

            for error in errors:
                report.add_error(error)

            for fix in fixes:
                report.add_fix(fix)

        return report

    def validate_and_highlight(self, ctx: DocumentContext) -> ValidationReport:
        report = ValidationReport()
        highlighter = DocumentHighlighter(ctx)

        for rule in self._rules:
            errors = rule.validate(ctx)
            for error in errors:
                report.add_error(error)

                hl = highlighter.highlight_error(error)
                if hl:
                    report.add_highlight(hl)

        return report

    def __repr__(self) -> str:
        rule_names = [r.name for r in self._rules]
        return f"<ValidationPipeline rules={rule_names}>"
