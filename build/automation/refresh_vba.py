"""Refresh the compiled VBA project without editing modules in the VBE."""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

import olefile
from pyopenvba import ExcelFile
from pyopenvba.exceptions import PyOpenVBAError

from ..paths import ROOT, VBA_BIN
from ..qa.excel import recalculate
from ..qa.vba_source import source_failures
from ..qa.verify_vba import check_vba
from ..vba.registry import (
    COMPILED_DOCUMENT_NAMES,
    SOURCE_BY_NAME,
    SOURCE_MODULE_NAMES,
    STANDARD_MODULE_NAMES,
)
from .workspace import ExcelWorkspaceError, excel_working_directory

if TYPE_CHECKING:
    from collections.abc import Sequence

LOGGER = logging.getLogger(__name__)
VBA_MEMBER = "xl/vbaProject.bin"
DEFAULT_SEED = ROOT / "PM_Workbook.xlsm"
EXPECTED_PROJECT_MODULES = (*COMPILED_DOCUMENT_NAMES, *STANDARD_MODULE_NAMES)
OLE_STREAM_PATH_PARTS = 2


class _RefreshProblem(Enum):
    SEED_MISSING = "VBA refresh seed does not exist: {}"
    SEED_FORMAT = "VBA refresh seed must be an .xlsm workbook: {}"
    SOURCE_GATE = "VBA source gate failed:\n{}"
    PACKAGE_OPEN = "cannot inspect {}: {}: {}"
    PACKAGE_DUPLICATE = "{} contains duplicate ZIP members: {}"
    PACKAGE_CORRUPT = "{} has a corrupt ZIP member: {}"
    PACKAGE_VBA_MISSING = "{} has no {}"
    PROJECT_OPEN = "cannot update the VBA project in {}: {}: {}"
    PROJECT_INVENTORY = "VBA project modules are {}; expected exactly {}"
    PROJECT_INVALID = "pyOpenVBA project validation failed:\n{}"
    PACKAGE_INVENTORY = "pyOpenVBA changed the ZIP member inventory or order"
    PACKAGE_NON_VBA = "pyOpenVBA changed non-VBA package members: {}"
    PACKAGE_VBA_UNCHANGED = "pyOpenVBA did not replace {}"
    SEED_COPY = "cannot copy VBA refresh seed {}: {}: {}"
    VBA_EXTRACT = "cannot extract {} from {}: {}: {}"
    SOURCE_ONLY_VERIFY = "source-only VBA verification failed:\n{}"
    EXCEL_COMPILE = "desktop Excel could not compile the disposable project: {}: {}"
    COMPILED_VERIFY = "compiled VBA verification failed:\n{}"
    CACHE_INSPECT = "cannot inspect compiled VBA caches in {}: {}: {}"
    PUBLISH = "cannot atomically publish {}: {}: {}"
    PUBLISH_VERIFY = "published VBA binary differs from the verified staged binary"
    PUBLISH_ROLLBACK = "VBA publication failed ({}) and rollback failed: {}: {}"
    WORKSPACE = "Excel automation workspace failed: {}"


class VbaRefreshError(RuntimeError):
    """An actionable failure in the source-to-compiled VBA refresh."""

    def __init__(self, problem: _RefreshProblem, *details: object) -> None:
        """Build the formatted refresh diagnostic."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True, slots=True)
class _ZipEntry:
    name: str
    date_time: tuple[int, int, int, int, int, int]
    compress_type: int
    comment: bytes
    extra: bytes
    create_system: int
    external_attr: int
    internal_attr: int
    flag_bits: int
    payload: bytes


@dataclass(frozen=True, slots=True)
class VbaRefreshResult:
    """Verified publication details for one refreshed VBA binary."""

    destination: Path
    source_modules: int
    cache_streams: int
    binary_bytes: int


def _diagnostics(problems: Sequence[str]) -> str:
    return "\n".join(f"  - {problem}" for problem in problems)


def _zip_snapshot(path: Path) -> tuple[_ZipEntry, ...]:
    try:
        package = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as error:
        raise VbaRefreshError(
            _RefreshProblem.PACKAGE_OPEN,
            path,
            type(error).__name__,
            error,
        ) from error

    with package:
        infos = package.infolist()
        counts = Counter(info.filename for info in infos)
        duplicates = sorted(name for name, count in counts.items() if count > 1)
        if duplicates:
            raise VbaRefreshError(
                _RefreshProblem.PACKAGE_DUPLICATE,
                path.name,
                ", ".join(duplicates),
            )
        corrupt = package.testzip()
        if corrupt is not None:
            raise VbaRefreshError(_RefreshProblem.PACKAGE_CORRUPT, path.name, corrupt)
        if VBA_MEMBER not in counts:
            raise VbaRefreshError(_RefreshProblem.PACKAGE_VBA_MISSING, path.name, VBA_MEMBER)
        return tuple(
            _ZipEntry(
                name=info.filename,
                date_time=info.date_time,
                compress_type=info.compress_type,
                comment=info.comment,
                extra=info.extra,
                create_system=info.create_system,
                external_attr=info.external_attr,
                internal_attr=info.internal_attr,
                flag_bits=info.flag_bits,
                payload=package.read(info),
            )
            for info in infos
        )


def _validate_project(book: ExcelFile) -> None:
    actual = tuple(book.module_names())
    if actual != EXPECTED_PROJECT_MODULES:
        raise VbaRefreshError(
            _RefreshProblem.PROJECT_INVENTORY,
            list(actual),
            list(EXPECTED_PROJECT_MODULES),
        )
    problems = book.validate()
    if problems:
        raise VbaRefreshError(_RefreshProblem.PROJECT_INVALID, _diagnostics(problems))


def _inject_sources(workbook: Path) -> None:
    try:
        with ExcelFile(workbook) as book:
            _validate_project(book)
            for module_name, source_path in SOURCE_BY_NAME.items():
                book.set_module(module_name, source_path.read_text(encoding="utf-8"))
            _validate_project(book)
            book.save()
        with ExcelFile(workbook) as saved_book:
            _validate_project(saved_book)
    except VbaRefreshError:
        raise
    except (KeyError, OSError, PyOpenVBAError, UnicodeError, ValueError) as error:
        raise VbaRefreshError(
            _RefreshProblem.PROJECT_OPEN,
            workbook.name,
            type(error).__name__,
            error,
        ) from error


def _package_change_problems(
    before: tuple[_ZipEntry, ...],
    after: tuple[_ZipEntry, ...],
) -> list[str]:
    before_names = tuple(entry.name for entry in before)
    after_names = tuple(entry.name for entry in after)
    if before_names != after_names:
        return [_RefreshProblem.PACKAGE_INVENTORY.value]

    changed_non_vba = [
        old.name
        for old, new in zip(before, after, strict=True)
        if old.name != VBA_MEMBER and old != new
    ]
    problems = []
    if changed_non_vba:
        problems.append(_RefreshProblem.PACKAGE_NON_VBA.value.format(", ".join(changed_non_vba)))
    old_vba = next(entry for entry in before if entry.name == VBA_MEMBER)
    new_vba = next(entry for entry in after if entry.name == VBA_MEMBER)
    if old_vba.payload == new_vba.payload:
        problems.append(_RefreshProblem.PACKAGE_VBA_UNCHANGED.value.format(VBA_MEMBER))
    return problems


def _extract_vba(workbook: Path, destination: Path) -> None:
    try:
        with zipfile.ZipFile(workbook) as package:
            destination.write_bytes(package.read(VBA_MEMBER))
    except (KeyError, OSError, zipfile.BadZipFile) as error:
        raise VbaRefreshError(
            _RefreshProblem.VBA_EXTRACT,
            VBA_MEMBER,
            workbook.name,
            type(error).__name__,
            error,
        ) from error


def _verify_binary(path: Path, *, require_compiled: bool) -> None:
    try:
        problems = check_vba(path, require_compiled=require_compiled)
    except (OSError, RuntimeError, ValueError) as error:
        problems = [f"{type(error).__name__}: {error}"]
    if problems:
        problem = (
            _RefreshProblem.COMPILED_VERIFY
            if require_compiled
            else _RefreshProblem.SOURCE_ONLY_VERIFY
        )
        raise VbaRefreshError(problem, _diagnostics(problems))


def _compile_in_excel(workbook: Path) -> None:
    try:
        recalculate(workbook)
    except (OSError, RuntimeError, ValueError) as error:
        raise VbaRefreshError(
            _RefreshProblem.EXCEL_COMPILE,
            type(error).__name__,
            error,
        ) from error


def _compiled_cache_count(path: Path) -> int:
    try:
        project = olefile.OleFileIO(str(path))
        try:
            return sum(
                1
                for stream in project.listdir(streams=True, storages=False)
                if len(stream) == OLE_STREAM_PATH_PARTS
                and stream[0] == "VBA"
                and stream[1].startswith("__SRP_")
            )
        finally:
            project.close()
    except (OSError, ValueError) as error:
        raise VbaRefreshError(
            _RefreshProblem.CACHE_INSPECT,
            path.name,
            type(error).__name__,
            error,
        ) from error


def _write_atomic(data: bytes, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, stage_name = tempfile.mkstemp(
        prefix=f".{destination.name}.", dir=destination.parent
    )
    stage = Path(stage_name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        stage.replace(destination)
    finally:
        active_error = sys.exception()
        try:
            if stage.exists():
                stage.unlink()
        except OSError as cleanup_error:
            if active_error is not None:
                active_error.add_note(
                    "staged binary cleanup also failed: "
                    f"{type(cleanup_error).__name__}: {cleanup_error}"
                )
            else:
                raise


def _restore_binary(destination: Path, original: bytes | None) -> None:
    if original is None:
        if destination.exists():
            destination.unlink()
    else:
        _write_atomic(original, destination)


def _require_published_bytes(destination: Path, expected: bytes) -> None:
    if destination.read_bytes() != expected:
        raise VbaRefreshError(_RefreshProblem.PUBLISH_VERIFY)


def _atomic_publish(source: Path, destination: Path) -> None:
    try:
        source_bytes = source.read_bytes()
        original = destination.read_bytes() if destination.exists() else None
    except OSError as error:
        raise VbaRefreshError(
            _RefreshProblem.PUBLISH,
            destination,
            type(error).__name__,
            error,
        ) from error

    try:
        _write_atomic(source_bytes, destination)
        _require_published_bytes(destination, source_bytes)
        _verify_binary(destination, require_compiled=True)
    except (OSError, VbaRefreshError) as publish_error:
        try:
            _restore_binary(destination, original)
        except OSError as rollback_error:
            raise VbaRefreshError(
                _RefreshProblem.PUBLISH_ROLLBACK,
                publish_error,
                type(rollback_error).__name__,
                rollback_error,
            ) from publish_error
        if isinstance(publish_error, VbaRefreshError):
            raise
        raise VbaRefreshError(
            _RefreshProblem.PUBLISH,
            destination,
            type(publish_error).__name__,
            publish_error,
        ) from publish_error


def _refresh_in_directory(seed: Path, destination: Path, work_dir: Path) -> VbaRefreshResult:
    source_problems = source_failures()
    if source_problems:
        raise VbaRefreshError(_RefreshProblem.SOURCE_GATE, _diagnostics(source_problems))

    disposable = work_dir / "PM_Workbook_VBA_Refresh.xlsm"
    source_only_bin = work_dir / "vbaProject.source-only.bin"
    compiled_bin = work_dir / "vbaProject.compiled.bin"
    try:
        shutil.copy2(seed, disposable)
    except OSError as error:
        raise VbaRefreshError(
            _RefreshProblem.SEED_COPY,
            seed,
            type(error).__name__,
            error,
        ) from error

    before = _zip_snapshot(disposable)
    _inject_sources(disposable)
    after = _zip_snapshot(disposable)
    package_problems = _package_change_problems(before, after)
    if package_problems:
        raise VbaRefreshError(_RefreshProblem.PROJECT_INVALID, _diagnostics(package_problems))

    _extract_vba(disposable, source_only_bin)
    _verify_binary(source_only_bin, require_compiled=False)
    _compile_in_excel(disposable)
    _extract_vba(disposable, compiled_bin)
    _verify_binary(compiled_bin, require_compiled=True)
    cache_streams = _compiled_cache_count(compiled_bin)
    _atomic_publish(compiled_bin, destination)
    return VbaRefreshResult(
        destination=destination,
        source_modules=len(SOURCE_MODULE_NAMES),
        cache_streams=cache_streams,
        binary_bytes=destination.stat().st_size,
    )


def refresh_vba(
    seed: str | Path = DEFAULT_SEED,
    destination: str | Path = VBA_BIN,
) -> VbaRefreshResult:
    """Inject current source, compile it in Excel, and publish the verified binary.

    Returns:
        Details of the atomically published compiled VBA project.

    Raises:
        VbaRefreshError: If source, package, Excel compilation, verification,
            publication, or cleanup fails.

    """
    seed_path = Path(seed).expanduser().resolve()
    destination_path = Path(destination).expanduser().resolve()
    if not seed_path.is_file():
        raise VbaRefreshError(_RefreshProblem.SEED_MISSING, seed_path)
    if seed_path.suffix.lower() != ".xlsm":
        raise VbaRefreshError(_RefreshProblem.SEED_FORMAT, seed_path)

    try:
        with excel_working_directory("pm-vba-refresh-") as work_dir:
            return _refresh_in_directory(seed_path, destination_path, work_dir)
    except ExcelWorkspaceError as error:
        raise VbaRefreshError(_RefreshProblem.WORKSPACE, error) from error


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Replace the complete VBA sources in a disposable workbook, compile them in "
            "desktop Excel, and atomically refresh build/vba/vbaProject.bin."
        )
    )
    parser.add_argument(
        "--seed",
        type=Path,
        default=DEFAULT_SEED,
        help="existing .xlsm package used only as the disposable VBA-project container",
    )
    parser.add_argument(
        "--destination",
        type=Path,
        default=VBA_BIN,
        help="compiled vbaProject.bin destination",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Run the guarded VBA refresh.

    Returns:
        Zero after verified publication, otherwise one.

    """
    arguments = _parser().parse_args(argv)
    failure: VbaRefreshError | None = None
    try:
        result = refresh_vba(arguments.seed, arguments.destination)
    except VbaRefreshError as error:
        failure = error
        result = None
    if failure is not None:
        notes = "".join(f"\n  - {note}" for note in getattr(failure, "__notes__", ()))
        LOGGER.error("VBA REFRESH FAILED - %s%s", failure, notes)
        return 1

    LOGGER.info(
        "VBA REFRESH PASS - %s source modules, %s compiled cache streams, %s bytes -> %s",
        result.source_modules,
        result.cache_streams,
        result.binary_bytes,
        result.destination,
    )
    return 0


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    raise SystemExit(main())
