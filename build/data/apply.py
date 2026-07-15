"""Approved provider-neutral workbook mutation orchestration."""

from __future__ import annotations

import hashlib
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from ..pipeline import PublicationPreconditionError, require_current_vba
from .bridge import PlanEvaluation, evaluate_plan
from .diagnostics import Diagnostic, DiagnosticPhase
from .inject import validate_snapshot
from .migrate import MigrationError, _require_workbook, rebuild_and_publish
from .snapshot import write_snapshot

if TYPE_CHECKING:
    from datetime import datetime

    from .snapshot import Snapshot

EXIT_SUCCESS = 0
EXIT_INVALID = 2
EXIT_CONFLICT = 3
EXIT_FAILURE = 4


@dataclass(frozen=True, kw_only=True, slots=True)
class ApplyEvaluation:
    """One JSON apply result paired with its documented process exit code."""

    result: dict[str, object]
    exit_code: int


def _diagnostic(
    code: str,
    phase: DiagnosticPhase,
    message: str,
    hint: str,
) -> Diagnostic:
    return Diagnostic(
        code=code,
        severity="error",
        phase=phase,
        pointer="",
        operation_id=None,
        message=message,
        hint=hint,
    )


def _error_text(error: Exception) -> str:
    details: list[str] = []
    current: BaseException | None = error
    seen: set[int] = set()
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        entry = f"{type(current).__name__}: {current}"
        notes = getattr(current, "__notes__", ())
        if notes:
            entry += " | notes: " + " | ".join(str(note) for note in notes)
        details.append(entry)
        current = current.__cause__ or current.__context__
    return " | caused by ".join(details)


def _apply_document(plan: PlanEvaluation) -> dict[str, object]:
    document = deepcopy(plan.result)
    document["result"] = "apply"
    document["applied"] = False
    document["publication"] = {"status": "not_started"}
    return document


def _append_error(
    document: dict[str, object],
    diagnostic: Diagnostic,
    *,
    conflict: bool,
) -> None:
    raw_errors = document.get("errors")
    errors = raw_errors if isinstance(raw_errors, list) else []
    errors.append(diagnostic.as_dict())
    document["errors"] = errors
    document["valid"] = False
    document["conflict"] = conflict


@dataclass(slots=True)
class _ApplyRunner:
    plan: PlanEvaluation
    source: Path
    approve: str

    def _plan_failure(self) -> ApplyEvaluation:
        document = _apply_document(self.plan)
        exit_code = EXIT_CONFLICT if self.plan.conflict else EXIT_INVALID
        return ApplyEvaluation(result=document, exit_code=exit_code)

    def _approval_failure(self) -> ApplyEvaluation:
        document = _apply_document(self.plan)
        diagnostic = _diagnostic(
            "approval.token_mismatch",
            "apply",
            "The supplied approval token does not match the current deterministic plan.",
            "Review the current plan token and approve that exact token before applying.",
        )
        _append_error(document, diagnostic, conflict=True)
        document["publication"] = {"status": "conflict"}
        return ApplyEvaluation(result=document, exit_code=EXIT_CONFLICT)

    def _workbook_failure(self, error: MigrationError) -> ApplyEvaluation:
        document = _apply_document(self.plan)
        diagnostic = _diagnostic(
            "apply.workbook",
            "apply",
            _error_text(error),
            "Supply the unlocked macro-enabled release workbook and retry.",
        )
        _append_error(document, diagnostic, conflict=False)
        return ApplyEvaluation(result=document, exit_code=EXIT_INVALID)

    def _no_change(self) -> ApplyEvaluation:
        document = _apply_document(self.plan)
        document["publication"] = {"status": "no_change", "workbook": str(self.source)}
        return ApplyEvaluation(result=document, exit_code=EXIT_SUCCESS)

    def _publication_conflict(
        self,
        error: PublicationPreconditionError,
        snapshot_path: Path,
    ) -> ApplyEvaluation:
        document = _apply_document(self.plan)
        diagnostic = _diagnostic(
            "approval.workbook_changed",
            "apply",
            _error_text(error),
            "Run describe and plan again against the workbook's new exact digest.",
        )
        _append_error(document, diagnostic, conflict=True)
        document["publication"] = {
            "status": "conflict",
            "workbook": str(self.source),
            "snapshot": str(snapshot_path),
        }
        return ApplyEvaluation(result=document, exit_code=EXIT_CONFLICT)

    def _publication_failure(
        self,
        error: Exception,
        snapshot_path: Path | None,
    ) -> ApplyEvaluation:
        document = _apply_document(self.plan)
        diagnostic = _diagnostic(
            "publication.failed",
            "publish",
            _error_text(error),
            "Resolve the build, Excel or publication failure; "
            "the source workbook was not approved.",
        )
        _append_error(document, diagnostic, conflict=False)
        publication: dict[str, object] = {
            "status": "failed",
            "workbook": str(self.source),
        }
        if snapshot_path is not None:
            publication["snapshot"] = str(snapshot_path)
        document["publication"] = publication
        return ApplyEvaluation(result=document, exit_code=EXIT_FAILURE)

    def _publish(self, base: Snapshot, intended: Snapshot) -> ApplyEvaluation:
        snapshot_path: Path | None = None
        try:
            require_current_vba()
            snapshot_path = write_snapshot(base)
            reconciliation = validate_snapshot(intended)
            backup = rebuild_and_publish(self.source, intended, reconciliation)
            workbook_digest = hashlib.sha256(self.source.read_bytes()).hexdigest()
        except PublicationPreconditionError as error:
            if snapshot_path is None:
                message = "publication precondition failed before snapshot persistence"
                raise RuntimeError(message) from error
            try:
                _require_workbook(self.source)
            except MigrationError as workbook_error:
                return self._workbook_failure(workbook_error)
            return self._publication_conflict(error, snapshot_path)
        except (
            BadZipFile,
            EOFError,
            InvalidFileException,
            OSError,
            RuntimeError,
            ValueError,
        ) as error:
            return self._publication_failure(error, snapshot_path)

        document = _apply_document(self.plan)
        document["applied"] = True
        document["publication"] = {
            "status": "published",
            "workbook": str(self.source),
            "snapshot": str(snapshot_path),
            "backup": str(backup),
            "workbook_sha256": workbook_digest,
        }
        return ApplyEvaluation(result=document, exit_code=EXIT_SUCCESS)

    def run(self) -> ApplyEvaluation:
        """Enforce approval before any persistence, build or publication.

        Returns:
            The apply document and its stable exit code.

        Raises:
            RuntimeError: If a valid plan has lost its complete snapshot state.

        """
        if not self.plan.valid:
            return self._plan_failure()
        if self.plan.token != self.approve:
            return self._approval_failure()
        try:
            self.source = _require_workbook(self.source)
        except MigrationError as error:
            return self._workbook_failure(error)
        if not self.plan.has_changes:
            return self._no_change()
        if self.plan.base_snapshot is None or self.plan.intended_snapshot is None:
            message = "valid changed plan has no complete snapshot state"
            raise RuntimeError(message)
        return self._publish(self.plan.base_snapshot, self.plan.intended_snapshot)


def evaluate_apply(
    payload: bytes,
    workbook: str | Path,
    *,
    approve: str,
    now: datetime | None = None,
) -> ApplyEvaluation:
    """Reparse, replan and conditionally publish one approved change set.

    Returns:
        The public apply result paired with exit code 0, 2, 3 or 4.

    """
    source = Path(workbook).expanduser().resolve()
    plan = evaluate_plan(payload, source, now=now)
    return _ApplyRunner(plan=plan, source=source, approve=approve).run()


def apply_workbook(
    payload: bytes,
    workbook: str | Path,
    *,
    approve: str,
    now: datetime | None = None,
) -> dict[str, object]:
    """Return one approved apply result.

    Returns:
        A JSON-compatible apply result.

    """
    return evaluate_apply(payload, workbook, approve=approve, now=now).result


__all__ = ["ApplyEvaluation", "apply_workbook", "evaluate_apply"]
