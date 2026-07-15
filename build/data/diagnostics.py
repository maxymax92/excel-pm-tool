"""Stable machine diagnostics for the provider-neutral data bridge."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

type DiagnosticSeverity = Literal["error", "warning"]
type DiagnosticPhase = Literal["parse", "schema", "describe", "plan", "apply", "publish"]


@dataclass(frozen=True, kw_only=True, slots=True)
class Diagnostic:
    """One stable, machine-readable bridge finding."""

    code: str
    severity: DiagnosticSeverity
    phase: DiagnosticPhase
    pointer: str
    operation_id: str | None
    message: str
    hint: str

    def as_dict(self) -> dict[str, str | None]:
        """Return the public JSON representation.

        Returns:
            A closed diagnostic object.

        """
        return {
            "code": self.code,
            "severity": self.severity,
            "phase": self.phase,
            "pointer": self.pointer,
            "operation_id": self.operation_id,
            "message": self.message,
            "hint": self.hint,
        }


class ContractError(ValueError):
    """Report one or more strict change-set parsing or schema failures."""

    def __init__(self, *diagnostics: Diagnostic) -> None:
        """Create an error from at least one structured diagnostic.

        Raises:
            ValueError: If no diagnostic is supplied.

        """
        if not diagnostics:
            message = "ContractError requires at least one diagnostic"
            raise ValueError(message)
        self.diagnostics = tuple(diagnostics)
        super().__init__(diagnostics[0].message)
