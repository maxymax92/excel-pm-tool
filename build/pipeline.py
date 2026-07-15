"""Build and transactionally publish the generated Excel workbooks.

The pipeline keeps authored packages separate from the disposable copies that
desktop Excel recalculates. Publication occurs only after exact authored
semantics survive both Excel rebuilds, and every destination is committed or
rolled back as one transaction.
"""

from __future__ import annotations

import logging
import os
import shutil
import sys
import tempfile
from dataclasses import dataclass
from datetime import UTC, datetime
from enum import Enum
from pathlib import Path
from stat import S_IMODE
from typing import TYPE_CHECKING
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

import xlsxwriter
from xlsxwriter.exceptions import XlsxWriterException

from .automation.workspace import excel_working_directory
from .core.design import MACRO_ACTIONS, WORKBOOK_WINDOW
from .core.formulas import encode_lambda
from .core.layout import REGISTRY
from .core.package_style import (
    PackageStyleError,
    PackageStyleResult,
    clone_zip_info,
    patch_workbook_package,
)
from .paths import DIST, ROOT, VBA_BIN
from .qa.excel import recalculate
from .qa.verify_excel import WorkbookSemanticError, compare_packages
from .qa.verify_vba import check_vba
from .spec import config
from .spec.capacity import CONFIG_ROWS, DATA_ROWS
from .spec.lambdas import LAMBDAS
from .writers import calc, data, views
from .writers.common import FONT, Formats

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping

    from xlsxwriter.workbook import Workbook

LOGGER = logging.getLogger(__name__)

# A constant document-created date makes identical inputs byte-deterministic.
BUILD_DATE = datetime(2026, 1, 1, tzinfo=UTC)
VBA_MEMBER = "xl/vbaProject.bin"

# Sheet order follows the daily workflow: stakeholder views, editable data,
# configuration, then the hidden calculation layer.
SHEETS = ["Overview", "Plan", "Items", "RAID", "Config", "Calc"]


class _PipelineProblem(Enum):
    VBA_DEPENDENCIES = "VBA verification requires the pinned oletools and olefile dependencies"
    VBA_MISMATCH = "VBA verification failed; no artifact was built:\n{}"
    VBA_EMBED = "XlsxWriter could not embed {}"
    VBA_MEMBER_COUNT = "{} contains {} {} members; expected exactly one"
    VBA_RESTORE_CRC = "{} has a CRC failure after VBA restoration: {}"
    VBA_RESTORE_MISMATCH = "{} does not embed the registered VBA binary after restoration"
    VBA_RESTORE_OPERATION = "cannot restore the registered VBA binary in {}: {}: {}"
    VBA_RESTORE_CLEANUP = "{} VBA restoration failed and temporary-file cleanup also failed: {}"
    ACTION_COUNT = "{} contains {} macro actions; expected {}"
    BUILD_CLEANUP = "{} build failed and temporary-file cleanup also failed: {}"
    DUPLICATE_DESTINATION = "publication destinations must be unique"
    IDENTICAL_SOURCE = "publication source and destination are identical: {}"
    NON_FILE_DESTINATION = "publication destination is not a file: {}"
    PREPARATION_CLEANUP = "release publication preparation failed and cleanup also failed: {}"
    ROLLBACK_CLEANUP = "rollback of {} failed and rollback-file cleanup also failed: {}"
    PUBLICATION_RECOVERED = "release publication failed; every destination was restored"
    PUBLICATION_DAMAGED = "release publication failed; {}"
    SUFFIX_MISMATCH = "semantic comparison requires matching workbook suffixes"
    SEMANTIC_CHANGE = (
        "desktop Excel changed authored workbook semantics; no artifact was published:\n{}"
    )


class BuildPipelineError(RuntimeError):
    """Base class for stable build and publication diagnostics."""

    def __init__(self, problem: _PipelineProblem, *details: object) -> None:
        """Create an error from a stable diagnostic template."""
        super().__init__(problem.value.format(*details))


class VbaVerificationError(BuildPipelineError):
    """Report missing or stale compiled VBA input."""


class WorkbookBuildError(BuildPipelineError):
    """Report workbook authoring, packaging, or cleanup failure."""


class PublicationError(BuildPipelineError):
    """Report a failed transactional publication or rollback."""


class SemanticPreservationError(BuildPipelineError):
    """Report authored workbook semantics changed by desktop Excel."""


class PublicationConfigurationError(ValueError):
    """Report an invalid publication plan before any destination changes."""

    def __init__(self, problem: _PipelineProblem, *details: object) -> None:
        """Create an error from a stable diagnostic template."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, slots=True)
class PublicationSnapshot:
    """Original bytes and mode for one publication destination."""

    data: bytes | None
    mode: int | None


@dataclass(frozen=True, slots=True)
class PublicationPlan:
    """Validated replacement and removal targets for one transaction."""

    replacements: tuple[tuple[Path, Path], ...]
    removals: tuple[Path, ...]

    @property
    def targets(self) -> tuple[Path, ...]:
        """Return every unique destination in transaction order.

        Returns:
            Replacement destinations followed by removal destinations.

        """
        return tuple(target for target, _source in self.replacements) + self.removals


_BUILD_EXCEPTIONS = (
    OSError,
    PackageStyleError,
    PublicationConfigurationError,
    VbaVerificationError,
    WorkbookBuildError,
    WorkbookSemanticError,
    XlsxWriterException,
)


def require_current_vba() -> None:
    """Require the compiled VBA project to match every source module.

    Raises:
        VbaVerificationError: If dependencies are missing or compiled code is stale.

    """
    try:
        problems = check_vba(VBA_BIN)
    except ImportError as error:
        raise VbaVerificationError(_PipelineProblem.VBA_DEPENDENCIES) from error
    if problems:
        details = "\n".join(f"  - {problem}" for problem in problems)
        raise VbaVerificationError(_PipelineProblem.VBA_MISMATCH, details)


def define_names(
    workbook: Workbook,
    setting_rows: Mapping[str, int],
    list_ranges: Mapping[str, str],
) -> None:
    """Define Config controls, validation sources, selectors, and LAMBDAs."""
    for name, _value, _description in config.SETTINGS:
        row = setting_rows[name] + 1
        workbook.define_name(name, f"=Config!$B${row}")

    for name, anchor in (
        ("lstActiveStatus", "Calc!$D$2"),
        ("lstDoneStatus", "Calc!$E$2"),
        ("lstCancelledStatus", "Calc!$F$2"),
        ("lstClosedRaid", "Calc!$G$2"),
        ("lstAlertRaid", "Calc!$H$2"),
        ("lstDecisionRaid", "Calc!$I$2"),
    ):
        workbook.define_name(name, f"=_xlfn.ANCHORARRAY({anchor})")

    for name, reference in (
        ("dvItemIDs", f"Calc!$A$2:$A${DATA_ROWS + 1}"),
        ("dvPeople", f"Calc!$B$2:$B${CONFIG_ROWS + 1}"),
        ("dvScopeLabels", f"Calc!$M$2:$M${DATA_ROWS + 2}"),
    ):
        workbook.define_name(name, f"={reference}")

    for name, reference in list_ranges.items():
        workbook.define_name(name, f"={reference}")

    for name, reference in (
        ("selPScope", "Plan!$B$2"),
        ("selPScope2", "Plan!$C$2"),
        ("selPScope3", "Plan!$C$3"),
        ("selPDepth", "Plan!$B$3"),
        ("selPFrom", "Plan!$E$2"),
        ("selPTo", "Plan!$E$3"),
        ("selPScopeID", "Plan!$BG$2"),
        ("selPScopeID2", "Plan!$BH$2"),
        ("selPScopeID3", "Plan!$BI$2"),
    ):
        workbook.define_name(name, f"={reference}")

    for name, parameters, body, extra in LAMBDAS:
        workbook.define_name(name, encode_lambda(parameters, body, extra))


def _temporary_output_path(destination: Path) -> Path:
    descriptor, name = tempfile.mkstemp(
        prefix=f".{destination.stem}.",
        suffix=destination.suffix,
        dir=destination.parent,
    )
    os.close(descriptor)
    temporary = Path(name)
    temporary.unlink()
    return temporary


def _new_workbook(path: Path) -> Workbook:
    workbook = xlsxwriter.Workbook(
        str(path),
        {
            "use_future_functions": False,
            "default_date_format": "dd mmm yyyy",
        },
    )
    workbook.formats[0].set_font_name(FONT["font_name"])
    workbook.formats[0].set_font_size(FONT["font_size"])
    workbook.set_properties({"created": BUILD_DATE})
    workbook.set_size(WORKBOOK_WINDOW["width"], WORKBOOK_WINDOW["height"])
    return workbook


def _embed_vba(workbook: Workbook) -> None:
    if workbook.add_vba_project(str(VBA_BIN)) != 0:
        raise WorkbookBuildError(_PipelineProblem.VBA_EMBED, VBA_BIN)
    workbook.set_vba_name("ThisWorkbook")


def _write_workbook(workbook: Workbook, *, with_vba: bool) -> None:
    REGISTRY.zones.clear()
    formats = Formats(workbook)
    worksheets = {name: workbook.add_worksheet(name) for name in SHEETS}

    views.write_overview(worksheets["Overview"], formats, is_xlsm=with_vba)
    views.write_plan(workbook, worksheets["Plan"], formats)
    data.write_items(worksheets["Items"], formats, is_xlsm=with_vba)
    data.write_raid(worksheets["RAID"], formats)
    setting_rows, list_ranges = data.write_config(worksheets["Config"], formats)
    calc.write_calc(worksheets["Calc"], formats)
    define_names(workbook, setting_rows, list_ranges)

    if with_vba:
        _embed_vba(workbook)
        for index, name in enumerate(SHEETS, start=1):
            worksheets[name].set_vba_name(f"Sheet{index}")

    worksheets["Overview"].activate()


def _require_action_count(
    destination: Path,
    styled: PackageStyleResult,
    *,
    with_vba: bool,
) -> None:
    expected = len(MACRO_ACTIONS) if with_vba else 0
    if styled.button_count != expected:
        raise WorkbookBuildError(
            _PipelineProblem.ACTION_COUNT,
            destination.name,
            styled.button_count,
            expected,
        )


def _cleanup_failed_build(workbook: Workbook | None, temporary: Path) -> list[BaseException]:
    errors: list[BaseException] = []
    if workbook is not None and not workbook.fileclosed:
        try:
            workbook.close()
        except (OSError, XlsxWriterException) as error:
            errors.append(error)
    if temporary.exists():
        try:
            temporary.unlink()
        except OSError as error:
            errors.append(error)
    return errors


def build_one(path: str | Path, *, with_vba: bool) -> PackageStyleResult:
    """Build one formula-only or macro-enabled workbook from source.

    Returns:
        The validated package-style result.

    Raises:
        WorkbookBuildError: If package creation or cleanup fails.

    """
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if with_vba:
        require_current_vba()

    temporary = _temporary_output_path(destination)
    workbook: Workbook | None = None
    try:
        workbook = _new_workbook(temporary)
        _write_workbook(workbook, with_vba=with_vba)
        workbook.close()
        workbook = None
        styled = patch_workbook_package(temporary)
        _require_action_count(destination, styled, with_vba=with_vba)
        temporary.replace(destination)
    except _BUILD_EXCEPTIONS as build_error:
        cleanup_errors = _cleanup_failed_build(workbook, temporary)
        if cleanup_errors:
            details = "; ".join(f"{type(error).__name__}: {error}" for error in cleanup_errors)
            raise WorkbookBuildError(
                _PipelineProblem.BUILD_CLEANUP,
                destination.name,
                details,
            ) from build_error
        raise

    LOGGER.info(
        "built %s (modern theme; %s macro actions)",
        destination.name,
        styled.button_count,
    )
    return styled


def _restore_snapshot(target: Path, snapshot: PublicationSnapshot) -> None:
    """Atomically restore one destination from its publication snapshot.

    Raises:
        PublicationError: If restoration and cleanup both fail.

    """
    if snapshot.data is None:
        if target.exists():
            target.unlink()
        return

    descriptor, name = tempfile.mkstemp(
        prefix=f".{target.stem}.rollback.",
        suffix=target.suffix,
        dir=target.parent,
    )
    restore_path = Path(name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(snapshot.data)
            stream.flush()
            os.fsync(stream.fileno())
        if snapshot.mode is not None:
            restore_path.chmod(snapshot.mode)
        restore_path.replace(target)
    except OSError as restore_error:
        cleanup_error: OSError | None = None
        if restore_path.exists():
            try:
                restore_path.unlink()
            except OSError as error:
                cleanup_error = error
        if cleanup_error is not None:
            detail = f"{type(cleanup_error).__name__}: {cleanup_error}"
            raise PublicationError(
                _PipelineProblem.ROLLBACK_CLEANUP,
                target,
                detail,
            ) from restore_error
        raise


def _cleanup_paths(paths: Iterable[Path]) -> list[OSError]:
    """Remove publication working files.

    Returns:
        Every cleanup failure, preserving traversal order.

    """
    errors: list[OSError] = []
    for path in paths:
        if not path.exists():
            continue
        try:
            path.unlink()
        except OSError as error:
            errors.append(error)
    return errors


def _write_registered_vba_archive(package: Path, temporary: Path, payload: bytes) -> None:
    with (
        ZipFile(package, "r") as source,
        ZipFile(temporary, "w", compression=ZIP_DEFLATED, allowZip64=True) as target,
    ):
        infos = source.infolist()
        member_count = sum(info.filename == VBA_MEMBER for info in infos)
        if member_count != 1:
            raise WorkbookBuildError(
                _PipelineProblem.VBA_MEMBER_COUNT,
                package.name,
                member_count,
                VBA_MEMBER,
            )
        target.comment = source.comment
        for info in infos:
            data = payload if info.filename == VBA_MEMBER else source.read(info)
            target.writestr(clone_zip_info(info), data)


def _validate_registered_vba_archive(package: Path, payload: bytes) -> None:
    with ZipFile(package, "r") as check:
        bad_part = check.testzip()
        if bad_part is not None:
            raise WorkbookBuildError(
                _PipelineProblem.VBA_RESTORE_CRC,
                package.name,
                bad_part,
            )
        if check.read(VBA_MEMBER) != payload:
            raise WorkbookBuildError(
                _PipelineProblem.VBA_RESTORE_MISMATCH,
                package.name,
            )


def _restore_registered_vba(package: Path) -> None:
    payload = VBA_BIN.read_bytes()
    original_mode = S_IMODE(package.stat().st_mode)
    temporary = _temporary_output_path(package)
    try:
        try:
            _write_registered_vba_archive(package, temporary, payload)
            _validate_registered_vba_archive(temporary, payload)
            temporary.chmod(original_mode)
            temporary.replace(package)
        except (BadZipFile, OSError) as error:
            raise WorkbookBuildError(
                _PipelineProblem.VBA_RESTORE_OPERATION,
                package.name,
                type(error).__name__,
                error,
            ) from error
    finally:
        operation_error = sys.exception()
        cleanup_errors = _cleanup_paths((temporary,))
        if cleanup_errors:
            details = "; ".join(f"{type(error).__name__}: {error}" for error in cleanup_errors)
            raise WorkbookBuildError(
                _PipelineProblem.VBA_RESTORE_CLEANUP,
                package.name,
                details,
            ) from operation_error


def _publication_plan(
    replacements: Mapping[str | Path, str | Path],
    removals: Iterable[str | Path],
) -> PublicationPlan:
    replacement_pairs = tuple(
        (Path(target), Path(source)) for target, source in replacements.items()
    )
    removal_targets = tuple(Path(target) for target in removals)
    plan = PublicationPlan(replacement_pairs, removal_targets)
    if len(plan.targets) != len(set(plan.targets)):
        raise PublicationConfigurationError(_PipelineProblem.DUPLICATE_DESTINATION)

    for target, source in plan.replacements:
        if not source.is_file():
            raise FileNotFoundError(source)
        target.parent.mkdir(parents=True, exist_ok=True)
        if source.resolve() == target.resolve():
            raise PublicationConfigurationError(_PipelineProblem.IDENTICAL_SOURCE, target)
    for target in plan.removals:
        target.parent.mkdir(parents=True, exist_ok=True)
    return plan


def _publication_snapshots(plan: PublicationPlan) -> dict[Path, PublicationSnapshot]:
    snapshots: dict[Path, PublicationSnapshot] = {}
    for target in plan.targets:
        if target.exists():
            if not target.is_file():
                raise PublicationConfigurationError(
                    _PipelineProblem.NON_FILE_DESTINATION,
                    target,
                )
            snapshots[target] = PublicationSnapshot(
                target.read_bytes(),
                S_IMODE(target.stat().st_mode),
            )
        else:
            snapshots[target] = PublicationSnapshot(None, None)
    return snapshots


def _prepare_publication(plan: PublicationPlan) -> dict[Path, Path]:
    prepared: dict[Path, Path] = {}
    try:
        for target, source in plan.replacements:
            descriptor, name = tempfile.mkstemp(
                prefix=f".{target.stem}.publish.",
                suffix=target.suffix,
                dir=target.parent,
            )
            os.close(descriptor)
            prepared_path = Path(name)
            prepared[target] = prepared_path
            shutil.copy2(source, prepared_path)
    except OSError as preparation_error:
        cleanup_errors = _cleanup_paths(prepared.values())
        if cleanup_errors:
            details = "; ".join(f"{type(error).__name__}: {error}" for error in cleanup_errors)
            raise PublicationError(
                _PipelineProblem.PREPARATION_CLEANUP,
                details,
            ) from preparation_error
        raise
    return prepared


def _apply_publication(plan: PublicationPlan, prepared: dict[Path, Path]) -> None:
    for target, _source in plan.replacements:
        prepared[target].replace(target)
        del prepared[target]
    for target in plan.removals:
        if target.exists():
            target.unlink()


def _restore_publication(
    plan: PublicationPlan,
    snapshots: Mapping[Path, PublicationSnapshot],
) -> list[BaseException]:
    errors: list[BaseException] = []
    for target in reversed(plan.targets):
        try:
            _restore_snapshot(target, snapshots[target])
        except (OSError, PublicationError) as error:
            errors.append(error)
    return errors


def _raise_publication_failure(
    operation_error: OSError,
    rollback_errors: Iterable[BaseException],
    cleanup_errors: Iterable[OSError],
) -> None:
    detail_parts: list[str] = []
    rollback_list = list(rollback_errors)
    cleanup_list = list(cleanup_errors)
    if rollback_list:
        detail_parts.append(
            "rollback failures: "
            + "; ".join(f"{type(error).__name__}: {error}" for error in rollback_list)
        )
    if cleanup_list:
        detail_parts.append(
            "cleanup failures: "
            + "; ".join(f"{type(error).__name__}: {error}" for error in cleanup_list)
        )
    if detail_parts:
        raise PublicationError(
            _PipelineProblem.PUBLICATION_DAMAGED,
            " | ".join(detail_parts),
        ) from operation_error
    raise PublicationError(_PipelineProblem.PUBLICATION_RECOVERED) from operation_error


def publish_transaction(
    replacements: Mapping[str | Path, str | Path],
    *,
    removals: Iterable[str | Path] = (),
) -> None:
    """Publish every release destination as one rollback-capable transaction."""
    plan = _publication_plan(replacements, removals)
    snapshots = _publication_snapshots(plan)
    prepared = _prepare_publication(plan)
    try:
        _apply_publication(plan, prepared)
    except OSError as publication_error:
        rollback_errors = _restore_publication(plan, snapshots)
        cleanup_errors = _cleanup_paths(prepared.values())
        _raise_publication_failure(publication_error, rollback_errors, cleanup_errors)


def _semantic_preservation_problems(source: Path, calculated: Path) -> list[str]:
    """Return labelled authored-semantic differences for one Excel rebuild.

    Returns:
        Semantic differences labelled with the calculated artifact name.

    Raises:
        PublicationConfigurationError: If workbook suffixes differ.

    """
    if source.suffix.lower() != calculated.suffix.lower():
        raise PublicationConfigurationError(_PipelineProblem.SUFFIX_MISMATCH)
    return [f"{calculated.name}: {issue}" for issue in compare_packages(source, calculated)]


def recalculate_stage(raw: Path, calculated: Path) -> None:
    """Copy one authored workbook and have desktop Excel recalculate the copy."""
    shutil.copy2(raw, calculated)
    recalculate(calculated)
    if calculated.suffix.lower() == ".xlsm":
        _restore_registered_vba(calculated)


def require_semantic_preservation(
    pairs: Iterable[tuple[Path, Path]],
) -> None:
    """Require every calculated copy to preserve its authored semantics.

    Raises:
        SemanticPreservationError: If desktop Excel changed authored semantics.

    """
    problems = [
        problem
        for source, calculated in pairs
        for problem in _semantic_preservation_problems(source, calculated)
    ]
    if problems:
        details = "\n".join(f"  - {problem}" for problem in problems)
        raise SemanticPreservationError(_PipelineProblem.SEMANTIC_CHANGE, details)


def _build_stage(stage: Path) -> tuple[Path, Path]:
    raw = stage / "raw"
    calculated = stage / "calculated"
    raw.mkdir()
    calculated.mkdir()

    raw_xlsx = raw / "PM_Workbook.xlsx"
    raw_xlsm = raw / "PM_Workbook.xlsm"
    calculated_xlsx = calculated / "PM_Workbook.xlsx"
    calculated_xlsm = calculated / "PM_Workbook.xlsm"

    build_one(raw_xlsx, with_vba=False)
    build_one(raw_xlsm, with_vba=True)
    recalculate_stage(raw_xlsx, calculated_xlsx)
    recalculate_stage(raw_xlsm, calculated_xlsm)
    require_semantic_preservation(((raw_xlsx, calculated_xlsx), (raw_xlsm, calculated_xlsm)))
    return calculated_xlsx, calculated_xlsm


def main() -> None:
    """Build, recalculate, verify, and transactionally publish both formats."""
    DIST.mkdir(exist_ok=True)
    require_current_vba()

    with excel_working_directory("pm-build-") as stage:
        calculated_xlsx, calculated_xlsm = _build_stage(stage)
        publish_transaction(
            {
                DIST / "PM_Workbook.xlsx": calculated_xlsx,
                DIST / "PM_Workbook.xlsm": calculated_xlsm,
                ROOT / "PM_Workbook.xlsm": calculated_xlsm,
            },
            removals=(ROOT / "PM_Workbook.xlsx",),
        )

    LOGGER.info(
        "published verified PM_Workbook.xlsm to the project root; "
        "the non-macro QA twin remains in dist/"
    )


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    main()
