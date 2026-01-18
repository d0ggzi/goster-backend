from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Iterator

from .report import ValidationError, AppliedFix

if TYPE_CHECKING:
    from .document import DocumentContext


class BaseRule(ABC):
    name: str = "base_rule"
    description: str = "Base validation rule"
    gost_reference: str = "ГОСТ 7.32-2017"

    @abstractmethod
    def validate(self, ctx: "DocumentContext") -> Iterator[ValidationError]:
        pass

    @abstractmethod
    def fix(self, ctx: "DocumentContext", error: ValidationError) -> AppliedFix | None:
        pass

    def validate_and_fix(
        self, ctx: "DocumentContext"
    ) -> tuple[list[ValidationError], list[AppliedFix]]:
        errors = list(self.validate(ctx))
        fixes = []

        for error in errors:
            if error.fixable:
                fix = self.fix(ctx, error)
                if fix:
                    fixes.append(fix)

        return errors, fixes

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}: {self.name}>"
