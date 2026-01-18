from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Iterable


class Severity(Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass
class ValidationError:
    rule_name: str
    message: str
    gost_reference: str
    severity: Severity = Severity.ERROR

    paragraph_index: int | None = None
    paragraph_text: str | None = None
    element_type: str | None = None

    current_value: Any | None = None
    expected_value: Any | None = None

    fixable: bool = True
    fix_data: dict = field(default_factory=dict)


@dataclass
class AppliedFix:
    rule_name: str
    error: ValidationError
    description: str
    before_value: Any | None = None
    after_value: Any | None = None


@dataclass
class AppliedHighlight:
    rule_name: str
    error: ValidationError
    paragraph_index: int | None
    color: str
    description: str


@dataclass
class ValidationReport:
    errors: list[ValidationError] = field(default_factory=list)
    fixes: list[AppliedFix] = field(default_factory=list)
    highlights: list[AppliedHighlight] = field(default_factory=list)

    @property
    def error_count(self) -> int:
        return len([e for e in self.errors if e.severity == Severity.ERROR])

    @property
    def warning_count(self) -> int:
        return len([e for e in self.errors if e.severity == Severity.WARNING])

    @property
    def fixed_count(self) -> int:
        return len(self.fixes)

    @property
    def highlighted_count(self) -> int:
        return len(self.highlights)

    @property
    def is_valid(self) -> bool:
        return self.error_count == 0

    def add_error(self, error: ValidationError) -> None:
        self.errors.append(error)

    def add_errors(self, errors: Iterable[ValidationError]) -> None:
        self.errors.extend(errors)

    def add_fix(self, fix: AppliedFix) -> None:
        self.fixes.append(fix)

    def add_highlight(self, highlight: AppliedHighlight) -> None:
        self.highlights.append(highlight)

    def summary(self) -> str:
        lines = [
            f"Найдено ошибок: {self.error_count}",
            f"Предупреждений: {self.warning_count}",
            f"Исправлено: {self.fixed_count}",
            f"Подсвечено: {self.highlighted_count}",
        ]
        return "\n".join(lines)

    def detailed_report(self) -> str:
        lines = [self.summary(), "", "=" * 50, ""]

        if self.errors:
            lines.append("ОШИБКИ:")
            lines.append("-" * 30)
            for i, err in enumerate(self.errors, 1):
                lines.append(f"{i}. [{err.severity.value.upper()}] {err.message}")
                lines.append(f"   Правило: {err.rule_name}")
                lines.append(f"   ГОСТ: {err.gost_reference}")
                if err.current_value is not None:
                    lines.append(f"   Текущее: {err.current_value}")
                if err.expected_value is not None:
                    lines.append(f"   Ожидается: {err.expected_value}")
                if err.paragraph_text:
                    preview = (
                        err.paragraph_text[:50] + "..."
                        if len(err.paragraph_text) > 50
                        else err.paragraph_text
                    )
                    lines.append(f'   Текст: "{preview}"')
                lines.append("")

        if self.fixes:
            lines.append("ИСПРАВЛЕНИЯ:")
            lines.append("-" * 30)
            for i, fix in enumerate(self.fixes, 1):
                lines.append(f"{i}. {fix.description}")
                lines.append(f"   Правило: {fix.rule_name}")
                if fix.before_value is not None:
                    lines.append(f"   Было: {fix.before_value}")
                if fix.after_value is not None:
                    lines.append(f"   Стало: {fix.after_value}")
                lines.append("")

        return "\n".join(lines)
