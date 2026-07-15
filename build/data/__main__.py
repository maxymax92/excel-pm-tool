"""Command-line entry point for the authored-data layer."""

import argparse
import logging
from pathlib import Path

from .migrate import DEFAULT_WORKBOOK, export_command, migrate_command
from .monday import AUTH_ENV, monday_command


def _add_workbook_argument(command: argparse.ArgumentParser) -> None:
    command.add_argument(
        "workbook",
        nargs="?",
        default=DEFAULT_WORKBOOK,
        type=Path,
        help=f"Workbook path (default: {DEFAULT_WORKBOOK.name} in the repository root).",
    )


def main() -> None:
    """Run one authored-data command."""
    parser = argparse.ArgumentParser(
        prog="build.data",
        description=(
            "Export authored workbook data, migrate a workbook onto rebuilt structure, "
            "or import a monday.com board."
        ),
    )
    commands = parser.add_subparsers(dest="command", required=True)
    for name, description in (
        ("export", "Export authored rows and settings into the snapshot ring."),
        ("migrate", "Rebuild the workbook from source and re-inject its own data."),
    ):
        command = commands.add_parser(name, help=description, description=description)
        _add_workbook_argument(command)

    monday_description = "Import one monday.com board into the Items hierarchy."
    monday = commands.add_parser("monday", help=monday_description, description=monday_description)
    monday.add_argument("board", type=int, help="monday.com board identifier.")
    _add_workbook_argument(monday)
    monday.add_argument(
        "--dry-run",
        action="store_true",
        help="Fetch, map and report without touching the workbook.",
    )
    monday.add_argument(
        "--allow-duplicates",
        action="store_true",
        help="Proceed although unmapped monday items share titles with existing rows.",
    )
    monday.add_argument(
        "--token-env",
        default=AUTH_ENV,
        help=f"Environment variable holding the API token (default: {AUTH_ENV}).",
    )

    arguments = parser.parse_args()
    if arguments.command == "export":
        export_command(arguments.workbook)
    elif arguments.command == "migrate":
        migrate_command(arguments.workbook)
    else:
        monday_command(
            arguments.board,
            arguments.workbook,
            dry_run=arguments.dry_run,
            allow_duplicates=arguments.allow_duplicates,
            token_env=arguments.token_env,
        )


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    main()
