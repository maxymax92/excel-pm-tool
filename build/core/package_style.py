"""Post-process generated Excel packages with the workbook design system.

XlsxWriter's default theme and DrawingML macro support do not match the release
contract. This module installs the Office theme, assigns macros to branded
DrawingML textboxes and rejects VML button controls after ``Workbook.close()``.
The rewrite is deliberately narrow: every other
package part, entry order and ZipInfo record is copied unchanged.
"""

from __future__ import annotations

import os
import re
import stat
import tempfile
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING, Literal
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from defusedxml import ElementTree as DefusedElementTree
from defusedxml.common import DefusedXmlException

if TYPE_CHECKING:
    from types import TracebackType

from ..paths import ASSETS
from .design import MACRO_ACTIONS, TYPOGRAPHY, WORKBOOK_WINDOW

THEME_PART = "xl/theme/theme1.xml"
WORKBOOK_PART = "xl/workbook.xml"
THEME_ASSET = ASSETS / "office-theme.xml"
DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Microsoft 365's current built-in Office Theme signature.  These are theme
# slots rather than workbook semantic colors, hence they live beside the
# package transformation and remain isolated from worksheet writers.
OFFICE_THEME_COLORS = {
    "dk2": "0E2841",
    "lt2": "E8E8E8",
    "accent1": "156082",
    "accent2": "E97132",
    "accent3": "196B24",
    "accent4": "0F9ED5",
    "accent5": "A02B93",
    "accent6": "4EA72E",
    "hlink": "467886",
    "folHlink": "96607D",
}

_VML_PATH = re.compile(r"^xl/drawings/[^/]+\.vml$", re.IGNORECASE)
_VML_SHAPE = re.compile(r"<v:shape\b.*?</v:shape>", re.IGNORECASE | re.DOTALL)
_DRAWING_PATH = re.compile(r"^xl/drawings/drawing\d+\.xml$", re.IGNORECASE)
_DRAWING_SHAPE = re.compile(r"<xdr:sp\b.*?</xdr:sp>", re.IGNORECASE | re.DOTALL)


class _PackageStyleProblem(Enum):
    ATTRIBUTE_TAG = "cannot add {!r} to a non-self-closing XML tag"
    PART_NOT_UTF8 = "{} is not UTF-8"
    MISSING_ELEMENT = "{} has no {} element"
    INVALID_XML = "invalid {}: {}"
    INVALID_TRANSFORMED_XML = "invalid transformed {}: {}"
    MISSING_THEME_SCHEME = "{} has no colour or font scheme"
    THEME_SLOT = "theme slot {} is {!r}; expected {!r}"
    THEME_FACE = "theme {} Latin face is {!r}; expected {!r}"
    EMPTY_PART = "{} is empty"
    THEME_ASSET = "cannot read theme asset {}: {}"
    UNIDENTIFIED_BUTTONS = "could not identify button shapes in {}"
    FORBIDDEN_BUTTONS = "{} contains {} forbidden VML button(s)"
    MISSING_SHAPE_OPEN = "action shape in {} has no xdr:sp opening tag"
    DUPLICATE_ACTION = "duplicate action shape in {}: {}"
    UNPATCHED_ACTION = "action shape was not patched in {}: {}"
    MISSING_PART = "missing required package part {}"
    DUPLICATE_DRAWING_ACTIONS = "duplicate action descriptions across drawing parts: {}"
    CRC_FAILURE = "CRC failure in rewritten part {}"
    CLEANUP_FAILURE = "package rewrite failed and temporary-file cleanup also failed: {}: {}"


class PackageStyleError(RuntimeError):
    """Report a generated package that violates the safe patch contract."""

    def __init__(self, problem: _PackageStyleProblem, *details: object) -> None:
        """Create a diagnostic from a stable problem template and its values."""
        super().__init__(problem.value.format(*details))


@dataclass(frozen=True)
class PackageStyleResult:
    """Small audit trail returned by :func:`patch_workbook_package`."""

    path: Path
    action_descriptions: tuple[str, ...]

    @property
    def button_count(self) -> int:
        """Return the number of DrawingML actions attached to shapes.

        Returns:
            The number of patched action descriptions.

        """
        return len(self.action_descriptions)


def _set_xml_attribute(tag: str, attribute: str, value: str) -> str:
    pattern = re.compile(rf'(\b{re.escape(attribute)}=")[^"]*(")')
    if pattern.search(tag):
        return pattern.sub(rf"\g<1>{value}\g<2>", tag, count=1)
    if not tag.endswith("/>"):
        raise PackageStyleError(_PackageStyleProblem.ATTRIBUTE_TAG, attribute)
    return tag[:-2] + f' {attribute}="{value}"/>'


def _patch_workbook_metadata(data: bytes, *, calculation_complete: bool) -> bytes:
    """Stop Excel treating the generated package as an Office 2007 file.

    XlsxWriter emits `lastEdited=4`, `lowestEdited=4` and the 2007
    `defaultThemeVersion`. Excel uses those compatibility markers to replace a
    perfectly valid modern theme with Office 2007 on the next save. Mark the
    workbook as a modern Excel document and let the explicit theme part govern.

    Returns:
        The transformed workbook metadata XML.

    Raises:
        PackageStyleError: If the source is malformed or misses a required element.

    """
    try:
        xml = data.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise PackageStyleError(_PackageStyleProblem.PART_NOT_UTF8, WORKBOOK_PART) from exc

    match = re.search(r"<fileVersion\b[^>]*/>", xml)
    if match is None:
        raise PackageStyleError(
            _PackageStyleProblem.MISSING_ELEMENT,
            WORKBOOK_PART,
            "fileVersion",
        )
    tag = match.group(0)
    for attribute, value in (("lastEdited", "7"), ("lowestEdited", "7"), ("rupBuild", "30203")):
        tag = _set_xml_attribute(tag, attribute, value)
    xml = xml[: match.start()] + tag + xml[match.end() :]

    pr = re.search(r"<workbookPr\b[^>]*/>", xml)
    if pr is None:
        raise PackageStyleError(_PackageStyleProblem.MISSING_ELEMENT, WORKBOOK_PART, "workbookPr")
    pr_tag = re.sub(r'\s+defaultThemeVersion="[^"]*"', "", pr.group(0))
    xml = xml[: pr.start()] + pr_tag + xml[pr.end() :]

    calc = re.search(r"<calcPr\b[^>]*/>", xml)
    if calc is None:
        raise PackageStyleError(_PackageStyleProblem.MISSING_ELEMENT, WORKBOOK_PART, "calcPr")
    calc_tag = calc.group(0)
    if calculation_complete:
        calc_tag = re.sub(r'\s+fullCalcOnLoad="[^"]*"', "", calc_tag)
        calc_tag = re.sub(r'\s+forceFullCalc="[^"]*"', "", calc_tag)
    xml = xml[: calc.start()] + calc_tag + xml[calc.end() :]

    # Store a full-screen-sized first-open window for the formula-only .xlsx.
    # SpreadsheetML dimensions are twips; XlsxWriter's public set_size() uses
    # pixels and applies the same 15x conversion. The macro workbook then uses
    # Workbook_Open/xlMaximized for an exact current-screen fit.
    view = re.search(r"<workbookView\b[^>]*/>", xml)
    if view is None:
        raise PackageStyleError(
            _PackageStyleProblem.MISSING_ELEMENT,
            WORKBOOK_PART,
            "workbookView",
        )
    view_tag = view.group(0)
    for attribute, value in (
        ("xWindow", str(WORKBOOK_WINDOW["x"])),
        ("yWindow", str(WORKBOOK_WINDOW["y"])),
        ("windowWidth", str(WORKBOOK_WINDOW["width"] * 15)),
        ("windowHeight", str(WORKBOOK_WINDOW["height"] * 15)),
        ("minimized", "0"),
    ):
        view_tag = _set_xml_attribute(view_tag, attribute, value)
    xml = xml[: view.start()] + view_tag + xml[view.end() :]

    result = xml.encode("utf-8")
    try:
        DefusedElementTree.fromstring(result)
    except (DefusedXmlException, SyntaxError) as exc:
        raise PackageStyleError(
            _PackageStyleProblem.INVALID_TRANSFORMED_XML,
            WORKBOOK_PART,
            exc,
        ) from exc
    return result


def _assert_theme_signature(data: bytes) -> None:
    """Prove the transformed theme is valid XML and has the expected values.

    Raises:
        PackageStyleError: If XML safety, colour or typography checks fail.

    """
    try:
        root = DefusedElementTree.fromstring(data)
    except (DefusedXmlException, SyntaxError) as exc:
        raise PackageStyleError(
            _PackageStyleProblem.INVALID_TRANSFORMED_XML,
            THEME_PART,
            exc,
        ) from exc

    ns = {"a": DRAWINGML_NS}
    scheme = root.find(".//a:clrScheme", ns)
    fonts = root.find(".//a:fontScheme", ns)
    if scheme is None or fonts is None:
        raise PackageStyleError(_PackageStyleProblem.MISSING_THEME_SCHEME, THEME_PART)

    for slot, expected in OFFICE_THEME_COLORS.items():
        node = scheme.find(f"a:{slot}/a:srgbClr", ns)
        actual = None if node is None else node.get("val")
        if actual != expected:
            raise PackageStyleError(_PackageStyleProblem.THEME_SLOT, slot, actual, expected)

    for family, expected in (
        ("major", TYPOGRAPHY["display_font"]),
        ("minor", TYPOGRAPHY["body_font"]),
    ):
        node = fonts.find(f"a:{family}Font/a:latin", ns)
        actual = None if node is None else node.get("typeface")
        if actual != expected:
            raise PackageStyleError(_PackageStyleProblem.THEME_FACE, family, actual, expected)


def _patch_theme(data: bytes) -> bytes:
    """Install the complete Office theme for round-trip stability.

    Returns:
        The validated Office theme asset.

    Raises:
        PackageStyleError: If the generated part or the theme asset is invalid.

    """
    if not data:
        raise PackageStyleError(_PackageStyleProblem.EMPTY_PART, THEME_PART)
    try:
        result = THEME_ASSET.read_bytes()
    except OSError as exc:
        raise PackageStyleError(_PackageStyleProblem.THEME_ASSET, THEME_ASSET, exc) from exc
    _assert_theme_signature(result)
    return result


def _is_button_shape(shape: str) -> bool:
    return bool(
        re.search(r'\bo:button\s*=\s*["\']t["\']', shape, re.IGNORECASE)
        or re.search(
            r'<x:ClientData\b[^>]*\bObjectType\s*=\s*["\']Button["\']',
            shape,
            re.IGNORECASE,
        )
    )


def _reject_vml_buttons(data: bytes, part_name: str) -> None:
    """Reject VML Form Control buttons; note-related VML may remain.

    Raises:
        PackageStyleError: If the part is malformed or contains a button control.

    """
    try:
        xml = data.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise PackageStyleError(_PackageStyleProblem.PART_NOT_UTF8, part_name) from exc

    try:
        DefusedElementTree.fromstring(data)
    except (DefusedXmlException, SyntaxError) as exc:
        raise PackageStyleError(_PackageStyleProblem.INVALID_XML, part_name, exc) from exc
    button_count = sum(1 for shape in _VML_SHAPE.findall(xml) if _is_button_shape(shape))
    if 'ObjectType="Button"' in xml and button_count == 0:
        raise PackageStyleError(_PackageStyleProblem.UNIDENTIFIED_BUTTONS, part_name)
    if button_count:
        raise PackageStyleError(_PackageStyleProblem.FORBIDDEN_BUTTONS, part_name, button_count)


def _patch_drawing_actions(data: bytes, part_name: str) -> tuple[bytes, set[str]]:
    """Attach VBA actions to the two branded DrawingML textboxes.

    XlsxWriter exposes deterministic textbox styling but emits ``macro=""``.
    Excel's SpreadsheetDrawing shape supports a macro name-reference, so the
    stable accessible description acts as the narrow patch key.

    Returns:
        The transformed XML and the descriptions patched in this drawing part.

    Raises:
        PackageStyleError: If an action shape is malformed, duplicated or missed.

    """
    try:
        xml = data.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise PackageStyleError(_PackageStyleProblem.PART_NOT_UTF8, part_name) from exc

    by_description = {action["description"]: action["macro"] for action in MACRO_ACTIONS.values()}
    seen: set[str] = set()

    def replace(shape_match: re.Match[str]) -> str:
        shape = shape_match.group(0)
        for description, macro in by_description.items():
            marker = re.compile(
                rf'<xdr:cNvPr\b[^>]*\bdescr="{re.escape(description)}"',
                re.IGNORECASE,
            )
            if not marker.search(shape):
                continue
            opening = re.search(r"<xdr:sp\b[^>]*>", shape, re.IGNORECASE)
            if opening is None:
                raise PackageStyleError(_PackageStyleProblem.MISSING_SHAPE_OPEN, part_name)
            patched = _set_xml_attribute(opening.group(0), "macro", f"[0]!{macro}")
            if description in seen:
                raise PackageStyleError(
                    _PackageStyleProblem.DUPLICATE_ACTION,
                    part_name,
                    description,
                )
            seen.add(description)
            return shape[: opening.start()] + patched + shape[opening.end() :]
        return shape

    updated = _DRAWING_SHAPE.sub(replace, xml)
    for description in by_description:
        if description in xml and description not in seen:
            raise PackageStyleError(
                _PackageStyleProblem.UNPATCHED_ACTION,
                part_name,
                description,
            )
    result = updated.encode("utf-8")
    try:
        DefusedElementTree.fromstring(result)
    except (DefusedXmlException, SyntaxError) as exc:
        raise PackageStyleError(
            _PackageStyleProblem.INVALID_TRANSFORMED_XML,
            part_name,
            exc,
        ) from exc
    return result, seen


def clone_zip_info(info: ZipInfo) -> ZipInfo:
    """Copy ZipInfo explicitly so reading and writing archives stay isolated.

    Returns:
        A detached copy of the package entry metadata.

    """
    clone = ZipInfo(info.filename, date_time=info.date_time)
    for attr in (
        "compress_type",
        "comment",
        "extra",
        "create_system",
        "create_version",
        "extract_version",
        "reserved",
        "flag_bits",
        "volume",
        "internal_attr",
        "external_attr",
    ):
        setattr(clone, attr, getattr(info, attr))
    return clone


class _TemporaryPackage:
    def __init__(self, package: Path) -> None:
        self.package = package
        self.path: Path | None = None

    def __enter__(self) -> Path:
        descriptor, name = tempfile.mkstemp(
            prefix=f".{self.package.name}.",
            suffix=".tmp",
            dir=self.package.parent,
        )
        os.close(descriptor)
        self.path = Path(name)
        return self.path

    def __exit__(
        self,
        _exc_type: type[BaseException] | None,
        error: BaseException | None,
        _traceback: TracebackType | None,
    ) -> Literal[False]:
        if self.path is None or not self.path.exists():
            return False
        try:
            self.path.unlink()
        except OSError as cleanup_error:
            if error is not None:
                raise PackageStyleError(
                    _PackageStyleProblem.CLEANUP_FAILURE,
                    type(cleanup_error).__name__,
                    cleanup_error,
                ) from error
            raise
        return False


def _transform_part(
    info: ZipInfo,
    data: bytes,
    action_descriptions: set[str],
    *,
    calculation_complete: bool,
) -> bytes:
    if info.filename == THEME_PART:
        return _patch_theme(data)
    if info.filename == WORKBOOK_PART:
        return _patch_workbook_metadata(data, calculation_complete=calculation_complete)
    if _DRAWING_PATH.match(info.filename):
        transformed, found = _patch_drawing_actions(data, info.filename)
        duplicates = action_descriptions & found
        if duplicates:
            descriptions = ", ".join(sorted(duplicates))
            raise PackageStyleError(
                _PackageStyleProblem.DUPLICATE_DRAWING_ACTIONS,
                descriptions,
            )
        action_descriptions.update(found)
        return transformed
    if _VML_PATH.match(info.filename):
        _reject_vml_buttons(data, info.filename)
    return data


def _rewrite_archive(
    package: Path,
    temp_path: Path,
    action_descriptions: set[str],
    *,
    calculation_complete: bool,
) -> None:
    with (
        ZipFile(package, "r") as source,
        ZipFile(temp_path, "w", compression=ZIP_DEFLATED, allowZip64=True) as target,
    ):
        target.comment = source.comment
        names = {info.filename for info in source.infolist()}
        for required_part in (THEME_PART, WORKBOOK_PART):
            if required_part not in names:
                raise PackageStyleError(_PackageStyleProblem.MISSING_PART, required_part)

        for info in source.infolist():
            data = _transform_part(
                info,
                source.read(info.filename),
                action_descriptions,
                calculation_complete=calculation_complete,
            )
            target.writestr(clone_zip_info(info), data)


def _validate_rewritten_archive(temp_path: Path) -> None:
    with ZipFile(temp_path, "r") as check:
        bad_part = check.testzip()
        if bad_part is not None:
            raise PackageStyleError(_PackageStyleProblem.CRC_FAILURE, bad_part)
        _assert_theme_signature(check.read(THEME_PART))


def patch_workbook_package(
    path: str | Path, *, calculation_complete: bool = False
) -> PackageStyleResult:
    """Patch an ``.xlsx`` or ``.xlsm`` OOXML package atomically in place.

    The original archive is replaced only after the rewritten temporary file
    passes Zip CRC validation and the theme signature check.  Entry order,
    timestamps, compression methods, permissions, comments and extra fields
    are retained wherever Python's standard ``zipfile`` API exposes them.
    Package contract violations propagate as :class:`PackageStyleError`.

    Returns:
        The patched package path and attached action descriptions.

    Raises:
        FileNotFoundError: If the requested package does not exist.

    """
    package = Path(path)
    if not package.is_file():
        raise FileNotFoundError(package)

    original_mode = stat.S_IMODE(package.stat().st_mode)
    action_descriptions: set[str] = set()
    with _TemporaryPackage(package) as temp_path:
        _rewrite_archive(
            package,
            temp_path,
            action_descriptions,
            calculation_complete=calculation_complete,
        )
        _validate_rewritten_archive(temp_path)
        temp_path.chmod(original_mode)
        temp_path.replace(package)

    return PackageStyleResult(
        path=package,
        action_descriptions=tuple(sorted(action_descriptions)),
    )


__all__ = [
    "PackageStyleError",
    "PackageStyleResult",
    "clone_zip_info",
    "patch_workbook_package",
]
