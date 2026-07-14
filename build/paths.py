"""Canonical repository and artifact paths used by the build package."""

from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
ROOT = PACKAGE_DIR.parent
DIST = ROOT / "dist"
ASSETS = PACKAGE_DIR / "assets"
AUTOMATION = PACKAGE_DIR / "automation"
VBA_DIR = PACKAGE_DIR / "vba"
VBA_BIN = VBA_DIR / "vbaProject.bin"

__all__ = [
    "ASSETS",
    "AUTOMATION",
    "DIST",
    "PACKAGE_DIR",
    "ROOT",
    "VBA_BIN",
    "VBA_DIR",
]
