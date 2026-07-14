"""Canonical VBA source and compiled-module inventory."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import TYPE_CHECKING

from ..paths import VBA_DIR

if TYPE_CHECKING:
    from pathlib import Path


class ModuleKind(Enum):
    """Supported VBA component kinds."""

    STANDARD = "standard"
    DOCUMENT = "document"


@dataclass(frozen=True, kw_only=True, slots=True)
class VbaModule:
    """One complete source-controlled VBA component."""

    name: str
    filename: str
    kind: ModuleKind
    private_to_project: bool
    public_procedures: tuple[str, ...]

    @property
    def path(self) -> Path:
        """Return the repository path for this module source.

        Returns:
            The complete VBA source path.

        """
        return VBA_DIR / self.filename


MODULES = (
    VbaModule(
        name="PMTool",
        filename="PMTool.bas",
        kind=ModuleKind.STANDARD,
        private_to_project=False,
        public_procedures=(
            "IsBlankValue",
            "ValidatedIdentifierText",
            "ExactTextCount",
            "ItemStatusRoles",
            "RaidStatusIsClosed",
            "IsBlockedDeliveryHealth",
            "NextUniqueIdInRange",
            "ExportMarkdown",
            "OrganiseItems",
        ),
    ),
    VbaModule(
        name="ThisWorkbook",
        filename="ThisWorkbook.cls.txt",
        kind=ModuleKind.DOCUMENT,
        private_to_project=False,
        public_procedures=(),
    ),
)

STANDARD_MODULES = tuple(module for module in MODULES if module.kind is ModuleKind.STANDARD)
DOCUMENT_MODULES = tuple(module for module in MODULES if module.kind is ModuleKind.DOCUMENT)
SOURCE_BY_NAME = {module.name: module.path for module in MODULES}
STANDARD_MODULE_NAMES = tuple(module.name for module in STANDARD_MODULES)
SOURCE_MODULE_NAMES = tuple(module.name for module in MODULES)
SOURCE_FILENAMES = frozenset(module.filename for module in MODULES)
HOST_SHEET_MODULE_NAMES = tuple(f"Sheet{index}" for index in range(1, 7))
COMPILED_DOCUMENT_NAMES = ("ThisWorkbook", *HOST_SHEET_MODULE_NAMES)
PUBLIC_MACROS = ("ExportMarkdown", "OrganiseItems")
WORKBOOK_EVENTS = ("Workbook_Open", "Workbook_SheetChange")
