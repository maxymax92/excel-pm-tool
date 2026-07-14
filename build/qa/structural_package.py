"""ZIP and OOXML package checks for generated workbooks."""

from __future__ import annotations

import zipfile
from collections import Counter
from typing import TYPE_CHECKING

from defusedxml import ElementTree as DefusedElementTree
from defusedxml.common import DefusedXmlException

from ..paths import VBA_BIN
from ..vba.registry import SOURCE_MODULE_NAMES
from .verify_vba import check_vba

if TYPE_CHECKING:
    from pathlib import Path


def _xml_member(name: str) -> bool:
    lowered = name.lower()
    return lowered.endswith((".xml", ".rels"))


def _xml_failures(package: zipfile.ZipFile) -> list[str]:
    failures: list[str] = []
    for member in package.infolist():
        if member.is_dir() or not _xml_member(member.filename):
            continue
        try:
            DefusedElementTree.fromstring(package.read(member))
        except (DefusedXmlException, SyntaxError) as error:
            failures.append(f"invalid OOXML part {member.filename}: {error}")
    return failures


def package_failures(path: Path) -> list[str]:
    """Check ZIP integrity, unique members, XML safety, and VBA presence.

    Returns:
        Every package-level contract violation.

    """
    try:
        package = zipfile.ZipFile(path)
    except zipfile.BadZipFile as error:
        return [f"invalid ZIP package: {error}"]

    with package:
        failures: list[str] = []
        member_counts = Counter(info.filename for info in package.infolist())
        duplicates = sorted(name for name, count in member_counts.items() if count > 1)
        failures.extend(f"duplicate ZIP member: {name}" for name in duplicates)

        bad_member = package.testzip()
        if bad_member is not None:
            failures.append(f"zip corrupt: {bad_member}")

        names = set(member_counts)
        has_vba = "xl/vbaProject.bin" in names
        if path.suffix.lower() == ".xlsm" and not has_vba:
            failures.append("vbaProject.bin missing from xlsm")
        if path.suffix.lower() == ".xlsx" and has_vba:
            failures.append("formula-only xlsx unexpectedly contains vbaProject.bin")
        if path.suffix.lower() == ".xlsm" and has_vba:
            embedded_vba = package.read("xl/vbaProject.bin")
            if not VBA_BIN.is_file():
                failures.append(f"compiled VBA input missing: {VBA_BIN}")
            elif embedded_vba != VBA_BIN.read_bytes():
                failures.append("embedded vbaProject.bin differs from the registered build input")
            failures.extend(
                f"VBA registry ({len(SOURCE_MODULE_NAMES)} modules): {problem}"
                for problem in check_vba(VBA_BIN)
            )
        failures.extend(_xml_failures(package))
        return failures
