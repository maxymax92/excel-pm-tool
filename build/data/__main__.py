"""Command-line entry point for authored data and the agent bridge."""

from __future__ import annotations

import argparse
import json
import logging
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, override
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .apply import EXIT_INVALID, ApplyEvaluation, evaluate_apply
from .bridge import describe_workbook, evaluate_plan
from .diagnostics import Diagnostic, DiagnosticPhase
from .migrate import DEFAULT_WORKBOOK, export_command, migrate_command
from .snapshot import atomic_write_json

if TYPE_CHECKING:
    from collections.abc import Sequence

LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True, slots=True)
class _CommandOutcome:
    document: dict[str, object]
    exit_code: int


class _CliInputError(ValueError):
    """Carry one structured command-line input diagnostic."""

    def __init__(self, diagnostic: Diagnostic) -> None:
        self.diagnostic = diagnostic
        super().__init__(diagnostic.message)


class _JsonArgumentParser(argparse.ArgumentParser):
    """Raise structured input errors instead of writing usage to stderr."""

    @override
    def error(self, message: str) -> None:
        """Raise one stable CLI argument diagnostic.

        Raises:
            _CliInputError: Always, instead of terminating the process directly.

        """
        raise _CliInputError(
            _diagnostic(
                "cli.arguments",
                "parse",
                message,
                "Correct the command arguments and retry.",
            )
        )


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


def _failure(result: str, diagnostic: Diagnostic) -> dict[str, object]:
    return {
        "result": result,
        "valid": False,
        "conflict": False,
        "warnings": [],
        "errors": [diagnostic.as_dict()],
    }


def _add_workbook_argument(command: argparse.ArgumentParser) -> None:
    command.add_argument(
        "workbook",
        nargs="?",
        default=DEFAULT_WORKBOOK,
        type=Path,
        help=f"Workbook path (default: {DEFAULT_WORKBOOK.name} in the repository root).",
    )


def _add_output_argument(command: argparse.ArgumentParser) -> None:
    command.add_argument(
        "--output",
        type=Path,
        help="Write the JSON result atomically to this file instead of stdout.",
    )


def _add_bridge_commands(commands: argparse._SubParsersAction[argparse.ArgumentParser]) -> None:
    describe = commands.add_parser(
        "describe",
        help="Describe the writable contract and current workbook state.",
    )
    _add_workbook_argument(describe)
    _add_output_argument(describe)

    for name, help_text in (
        ("plan", "Validate and diff one provider-neutral change set."),
        ("apply", "Replan and publish one explicitly approved change set."),
    ):
        command = commands.add_parser(name, help=help_text, description=help_text)
        command.add_argument("change_set", help="UTF-8 JSON file, or '-' for stdin.")
        _add_workbook_argument(command)
        _add_output_argument(command)
        if name == "apply":
            command.add_argument(
                "--approve",
                required=True,
                help="Exact plan token reviewed by the user.",
            )


def _parser() -> argparse.ArgumentParser:
    parser = _JsonArgumentParser(
        prog="build.data",
        description=(
            "Export or migrate authored workbook data, or describe, plan and apply "
            "provider-neutral agent change sets."
        ),
    )
    commands = parser.add_subparsers(
        dest="command",
        required=True,
        parser_class=_JsonArgumentParser,
    )
    for name, description in (
        ("export", "Export authored rows and settings into the snapshot ring."),
        ("migrate", "Rebuild the workbook from source and re-inject its own data."),
    ):
        command = commands.add_parser(name, help=description, description=description)
        _add_workbook_argument(command)
    _add_bridge_commands(commands)
    return parser


def _output_path(arguments: argparse.Namespace) -> Path | None:
    raw_output = getattr(arguments, "output", None)
    if raw_output is None:
        return None
    output = Path(raw_output).expanduser().resolve()
    workbook = Path(arguments.workbook).expanduser().resolve()
    forbidden = {workbook}
    change_set = getattr(arguments, "change_set", "-")
    if change_set != "-":
        forbidden.add(Path(change_set).expanduser().resolve())
    if output in forbidden:
        raise _CliInputError(
            _diagnostic(
                "cli.output_collision",
                "parse",
                f"Output path {output} aliases a protected input file.",
                "Choose an output path different from the workbook and change-set input.",
            )
        )
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists() and not output.is_file():
            raise IsADirectoryError(output)
        with tempfile.TemporaryFile(dir=output.parent):
            pass
    except OSError as error:
        raise _CliInputError(
            _diagnostic(
                "cli.output_unwritable",
                "parse",
                f"{type(error).__name__}: {error}",
                "Choose a writable output path and retry.",
            )
        ) from error
    return output


def _read_change_set(value: str) -> bytes:
    if value == "-":
        return sys.stdin.buffer.read()
    return Path(value).expanduser().resolve().read_bytes()


def _run_describe(arguments: argparse.Namespace) -> _CommandOutcome:
    try:
        document = describe_workbook(arguments.workbook)
    except (BadZipFile, EOFError, InvalidFileException, OSError, RuntimeError, ValueError) as error:
        diagnostic = _diagnostic(
            "workbook.unreadable",
            "describe",
            f"{type(error).__name__}: {error}",
            "Supply an existing, structurally valid workbook and retry.",
        )
        return _CommandOutcome(_failure("describe", diagnostic), EXIT_INVALID)
    return _CommandOutcome(document, 0)


def _run_plan(arguments: argparse.Namespace, payload: bytes) -> _CommandOutcome:
    evaluation = evaluate_plan(payload, arguments.workbook)
    exit_code = 3 if evaluation.conflict else 0 if evaluation.valid else EXIT_INVALID
    return _CommandOutcome(evaluation.result, exit_code)


def _run_apply(arguments: argparse.Namespace, payload: bytes) -> _CommandOutcome:
    evaluation: ApplyEvaluation = evaluate_apply(
        payload,
        arguments.workbook,
        approve=arguments.approve,
    )
    return _CommandOutcome(evaluation.result, evaluation.exit_code)


def _run_bridge(arguments: argparse.Namespace) -> _CommandOutcome:
    if arguments.command == "describe":
        return _run_describe(arguments)
    try:
        payload = _read_change_set(arguments.change_set)
    except OSError as error:
        diagnostic = _diagnostic(
            "cli.change_set_read",
            "parse",
            f"{type(error).__name__}: {error}",
            "Supply a readable UTF-8 JSON change-set file or '-' for stdin.",
        )
        return _CommandOutcome(_failure(arguments.command, diagnostic), EXIT_INVALID)
    if arguments.command == "plan":
        return _run_plan(arguments, payload)
    return _run_apply(arguments, payload)


def _emit(document: dict[str, object], output: Path | None) -> None:
    if output is None:
        json.dump(document, sys.stdout, ensure_ascii=False, indent=2, sort_keys=True)
        sys.stdout.write("\n")
        return
    output.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_json(output, document)
    LOGGER.info("result: %s", output)


def _run_existing(arguments: argparse.Namespace) -> int:
    if arguments.command == "export":
        export_command(arguments.workbook)
    else:
        migrate_command(arguments.workbook)
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    """Run one authored-data command.

    Returns:
        The documented process exit code.

    """
    try:
        arguments = _parser().parse_args(argv)
    except _CliInputError as error:
        _emit(_failure("cli", error.diagnostic), None)
        return EXIT_INVALID
    if arguments.command in {"export", "migrate"}:
        return _run_existing(arguments)
    try:
        output = _output_path(arguments)
    except _CliInputError as error:
        _emit(_failure(arguments.command, error.diagnostic), None)
        return EXIT_INVALID

    outcome = _run_bridge(arguments)
    try:
        _emit(outcome.document, output)
    except OSError as error:
        diagnostic = _diagnostic(
            "cli.output_write",
            "publish",
            f"{type(error).__name__}: {error}",
            "Inspect workbook state before retrying apply, then choose a writable output path.",
        )
        _emit(_failure(arguments.command, diagnostic), None)
        return 4 if arguments.command == "apply" else EXIT_INVALID
    return outcome.exit_code


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    raise SystemExit(main())
