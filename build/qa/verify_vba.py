"""Verify the compiled VBA project against the complete source registry."""

from __future__ import annotations

import logging
import re
import sys
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

import olefile
from oletools.olevba import VBA_Parser

from ..paths import VBA_BIN
from ..vba.registry import (
    COMPILED_DOCUMENT_NAMES,
    SOURCE_BY_NAME,
    SOURCE_MODULE_NAMES,
    STANDARD_MODULE_NAMES,
)

if TYPE_CHECKING:
    from collections.abc import Sequence

LOGGER = logging.getLogger(__name__)
VB_NAME_RE = re.compile(r'(?im)^Attribute VB_Name = "([^"]+)"\s*$')


class _VbaProblem(Enum):
    USAGE = "usage: python -m build.qa.verify_vba [build/vba/vbaProject.bin]"
    PROJECT_NON_ASCII = "PROJECT standard-module name is not ASCII: {!r}"
    COMPILED_CACHE_MISSING = (
        "compiled VBA cache streams are missing; desktop Excel must compile and save the project"
    )
    COMPILED_CACHE_EMPTY = (
        "VBA/_VBA_PROJECT contains no compiled cache body; desktop Excel must compile and save "
        "the project"
    )


class _VbaUsageError(ValueError):
    def __init__(self) -> None:
        super().__init__(_VbaProblem.USAGE.value)


@dataclass(frozen=True, slots=True)
class _ExtractedModule:
    container_name: str
    code: str
    declared_names: tuple[str, ...]


def _norm(code: str) -> str:
    """Remove VBE metadata without changing behavioral source text.

    Returns:
        Source with package metadata and permitted whitespace differences removed.

    """
    lines = [
        line.rstrip(" \t")
        for line in code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if not line.lstrip().startswith("Attribute ")
    ]
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def _normalizer_contract_problems() -> list[str]:
    """Protect exact comparison from lossy normalization.

    Returns:
        Diagnostics when permitted or behavioral differences are mishandled.

    """
    equivalent_pairs = (
        ("Option Explicit\r\n", "Option Explicit\n"),
        ('Attribute VB_Name = "PMTool"\nOption Explicit', "Option Explicit"),
        ("Option Explicit   \n", "Option Explicit\n"),
    )
    distinct_pairs = (
        ('MsgBox "Blocked item"', 'MsgBox "blocked item"'),
        ('MsgBox "two  spaces"', 'MsgBox "two spaces"'),
        ("' Current behavior", "' current behavior"),
        ("    value = 1", "value = 1"),
        ("Option Explicit\n\nPrivate Sub A()", "Option Explicit\nPrivate Sub A()"),
    )
    problems = [
        "VBA normalizer does not ignore permitted package differences"
        for left, right in equivalent_pairs
        if _norm(left) != _norm(right)
    ]
    problems.extend(
        "VBA normalizer hides a behavioral source difference"
        for left, right in distinct_pairs
        if _norm(left) == _norm(right)
    )
    return problems


def _project_modules(bin_path: Path, problems: list[str]) -> list[str]:
    project = olefile.OleFileIO(str(bin_path))
    try:
        data = project.openstream("PROJECT").read()
    finally:
        project.close()

    names: list[str] = []
    for line in data.replace(b"\r\n", b"\n").replace(b"\r", b"\n").split(b"\n"):
        if not line.startswith(b"Module="):
            continue
        raw_name = line.split(b"=", 1)[1].strip()
        try:
            names.append(raw_name.decode("ascii"))
        except UnicodeDecodeError:
            problems.append(_VbaProblem.PROJECT_NON_ASCII.value.format(raw_name))
    return names


def _extract_modules(bin_path: Path, problems: list[str]) -> list[_ExtractedModule]:
    parser = VBA_Parser(str(bin_path))
    modules: list[_ExtractedModule] = []
    try:
        for _filename, _stream_path, container_name, code in parser.extract_macros():
            source = code or ""
            modules.append(
                _ExtractedModule(
                    container_name,
                    source,
                    tuple(VB_NAME_RE.findall(source)),
                )
            )
    finally:
        parser.close()

    container_names = [module.container_name for module in modules]
    problems.extend(
        f"duplicate embedded VBA container name: {name!r}"
        for name in sorted({name for name in container_names if container_names.count(name) > 1})
    )
    problems.extend(
        f"module {module.container_name!r} has {len(module.declared_names)} "
        "Attribute VB_Name lines; expected at most 1"
        for module in modules
        if len(module.declared_names) > 1
    )
    return modules


def _check_project_registration(bin_path: Path, problems: list[str]) -> None:
    project_modules = _project_modules(bin_path, problems)
    if len(project_modules) != len(set(project_modules)):
        problems.append(f"duplicate PROJECT module entries: {project_modules!r}")
    if tuple(project_modules) != STANDARD_MODULE_NAMES:
        displayed = project_modules or ["<none>"]
        problems.append(
            f"standard modules are registered as {displayed!r}; expected exactly "
            f"{list(STANDARD_MODULE_NAMES)!r} in registry order"
        )


def _check_compiled_cache(bin_path: Path, problems: list[str]) -> None:
    project = olefile.OleFileIO(str(bin_path))
    try:
        stream_names = {"/".join(path) for path in project.listdir(streams=True, storages=False)}
        cache_streams = {name for name in stream_names if name.startswith("VBA/__SRP_")}
        cache_body = project.openstream("VBA/_VBA_PROJECT").read()[7:]
    finally:
        project.close()

    if not cache_streams:
        problems.append(_VbaProblem.COMPILED_CACHE_MISSING.value)
    if not cache_body or not any(cache_body):
        problems.append(_VbaProblem.COMPILED_CACHE_EMPTY.value)


def _executable_source(code: str) -> str:
    return "\n".join(
        line
        for line in _norm(code).splitlines()
        if line and line != "Option Explicit" and not line.lstrip().startswith("'")
    )


def _declared_module_map(
    modules: list[_ExtractedModule],
    problems: list[str],
) -> dict[str, _ExtractedModule]:
    declared: dict[str, _ExtractedModule] = {}
    for module in modules:
        if not module.declared_names:
            if _executable_source(module.code):
                problems.append(
                    f"unexpected executable VBA without VB_Name remains in "
                    f"{module.container_name!r}"
                )
            continue
        if len(module.declared_names) != 1:
            continue
        name = module.declared_names[0]
        if name in declared:
            problems.append(f"duplicate embedded VB_Name: {name!r}")
            continue
        declared[name] = module
    return declared


def _first_difference(actual: str, expected: str) -> str:
    actual_lines = actual.split("\n")
    expected_lines = expected.split("\n")
    first = next(
        (
            index
            for index in range(max(len(actual_lines), len(expected_lines)))
            if (actual_lines[index] if index < len(actual_lines) else None)
            != (expected_lines[index] if index < len(expected_lines) else None)
        ),
        None,
    )
    if first is None:
        return ""
    actual_text = actual_lines[first] if first < len(actual_lines) else "<end>"
    expected_text = expected_lines[first] if first < len(expected_lines) else "<end>"
    return f" first diff ~line {first + 1}: shipped {actual_text!r} vs source {expected_text!r}"


def _check_required_sources(
    declared: dict[str, _ExtractedModule],
    problems: list[str],
) -> None:
    for source_name, source_path in SOURCE_BY_NAME.items():
        module = declared.get(source_name)
        if module is None:
            problems.append(f"shipped bin has no {source_name} module")
            continue
        if not source_path.is_file():
            problems.append(f"VBA source file is missing: {source_path}")
            continue
        actual = _norm(module.code)
        expected = _norm(source_path.read_text(encoding="utf-8"))
        if actual != expected:
            problems.append(
                f"{source_name} compiled code differs from build/vba/{source_path.name}."
                + _first_difference(actual, expected)
            )


def _check_exact_inventory(
    declared: dict[str, _ExtractedModule],
    problems: list[str],
) -> None:
    allowed = set(SOURCE_MODULE_NAMES) | set(COMPILED_DOCUMENT_NAMES)
    actual = set(declared)
    problems.extend(f"compiled VBA module missing: {name}" for name in sorted(allowed - actual))
    problems.extend(f"unexpected compiled VBA module: {name}" for name in sorted(actual - allowed))
    for sheet_name in COMPILED_DOCUMENT_NAMES[1:]:
        module = declared.get(sheet_name)
        if module is not None and _executable_source(module.code):
            problems.append(f"host document module {sheet_name!r} contains executable code")


def check_vba(
    bin_path: str | Path,
    *,
    require_compiled: bool = True,
) -> list[str]:
    """Compare compiled registration, inventory, and code with all sources.

    Returns:
        Every registration, inventory, source, and package mismatch.

    """
    package_path = Path(bin_path)
    problems = _normalizer_contract_problems()
    if not package_path.is_file():
        return [*problems, f"no vbaProject.bin at {package_path}"]

    if require_compiled:
        _check_compiled_cache(package_path, problems)
    modules = _extract_modules(package_path, problems)
    _check_project_registration(package_path, problems)
    declared = _declared_module_map(modules, problems)
    _check_exact_inventory(declared, problems)
    _check_required_sources(declared, problems)
    return problems


def main(argv: Sequence[str] | None = None) -> int:
    """Report whether compiled VBA exactly matches the full source registry.

    Returns:
        A process exit status: zero for success and one for mismatches.

    Raises:
        _VbaUsageError: If more than one optional binary path is supplied.

    """
    arguments = sys.argv[1:] if argv is None else argv
    if len(arguments) > 1:
        raise _VbaUsageError
    bin_path = Path(arguments[0]) if arguments else VBA_BIN
    problems = check_vba(bin_path)
    if problems:
        LOGGER.error("VBA VERIFICATION FAILED - %s problem(s):", len(problems))
        for problem in problems:
            LOGGER.error("  - %s", problem)
        LOGGER.error("\nSee docs/vba-maintenance.md for the compilation procedure.")
        return 1
    LOGGER.info(
        "VBA VERIFICATION PASS - %s matches all %s source modules",
        bin_path.name,
        len(SOURCE_MODULE_NAMES),
    )
    return 0


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    raise SystemExit(main())
