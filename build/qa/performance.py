"""Measure interactive workbook latency in desktop Excel for Mac.

The benchmark opens the workbook without saving, switches among visible tabs,
selects ordinary Items cells, and edits/restores Latest Status and Delivery
Health. It also builds and measures a populated 500-row formula workbook. Gross
latency regressions and calculation-mode restoration failures are hard failures.

Usage:
    uv run --frozen python -m build.qa.performance workbook
"""

from __future__ import annotations

import asyncio
import re
import sys
from datetime import UTC, datetime, timedelta
from pathlib import Path

from ..automation.workspace import excel_workbook_copy
from ..paths import AUTOMATION
from ..pipeline import build_one
from .common import temporary_examples, temporary_workbook
from .excel import recalculate

LIMITS_MS = {
    "OPEN_MS": 2500.0,
    "TAB_SWITCH_MS": 100.0,
    "CELL_SELECT_MS": 50.0,
    "EDIT_MS": 150.0,
    "CALC_EDIT_MS": 200.0,
}
SCALE_ROWS = 500
SCALE_LIMITS_MS = {
    "OPEN_MS": 6000.0,
    "TAB_SWITCH_MS": 150.0,
    "CELL_SELECT_MS": 75.0,
    "EDIT_MS": 250.0,
    "CALC_EDIT_MS": 500.0,
}
OPEN_RETRY_LIMIT = 30
EXPECTED_METRICS = set(LIMITS_MS) | {"OPEN_RETRIES", "CALCULATION_RESTORED"}
OSASCRIPT = Path("/usr/bin/osascript")
MIN_SCALE_ROWS = 2
EXPECTED_ARG_COUNT = 2


def benchmark_contract_failures() -> list[str]:
    """Check calculation-mode capture, force and restoration on both exits.

    Returns:
        Every static contract failure found in the benchmark script.

    """
    source_path = AUTOMATION / "excel_benchmark.applescript"
    source = source_path.read_text(encoding="utf-8")
    error_marker = "    on error failureMessage number failureNumber"
    if source.count(error_marker) != 1:
        return ["Excel benchmark has no unambiguous failure handler"]
    success_path, failure_path = source.split(error_marker, 1)
    failures = [
        f"Excel benchmark success path is missing {token}"
        for token in (
            "set previousCalculation to calculation",
            "set calculation to calculation automatic",
            "set calculation to previousCalculation",
            "if calculation is not previousCalculation then",
        )
        if token not in success_path
    ]
    failures.extend(
        f"Excel benchmark failure path is missing {token}"
        for token in (
            "set calculation to previousCalculation",
            "calculation restore failed:",
        )
        if token not in failure_path
    )
    return failures


async def _run_benchmark_command(
    script: Path,
    workbook: Path,
    timeout_seconds: int,
) -> tuple[int, str, str]:
    process = await asyncio.create_subprocess_exec(
        str(OSASCRIPT),
        str(script),
        str(workbook),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    try:
        async with asyncio.timeout(timeout_seconds):
            stdout, stderr = await process.communicate()
    except TimeoutError as error:
        process.kill()
        await process.wait()
        message = f"Excel benchmark exceeded its {timeout_seconds}-second timeout"
        raise TimeoutError(message) from error
    return process.returncode or 0, stdout.decode(), stderr.decode()


def measure(path: Path, *, timeout: int = 180) -> dict[str, float]:
    """Return desktop-Excel latency metrics for one workbook.

    Returns:
        Parsed latency and state-restoration metrics.

    Raises:
        FileNotFoundError: If the workbook, benchmark script, or osascript is missing.
        RuntimeError: If the benchmark process or result contract fails.

    """
    workbook = path.expanduser().resolve()
    if not workbook.is_file():
        raise FileNotFoundError(workbook)
    script = AUTOMATION / "excel_benchmark.applescript"
    if not script.is_file():
        raise FileNotFoundError(script)
    if not OSASCRIPT.is_file():
        raise FileNotFoundError(OSASCRIPT)
    with excel_workbook_copy(workbook, prefix=f"{workbook.stem}.benchmark.") as disposable:
        returncode, stdout_raw, stderr_raw = asyncio.run(
            _run_benchmark_command(script, disposable, timeout)
        )
    stdout = stdout_raw.strip()
    stderr = stderr_raw.strip()
    if returncode != 0:
        msg = (
            f"Excel benchmark failed for {workbook.name} "
            f"(exit {returncode}): {stderr or stdout or 'no diagnostic'}"
        )
        raise RuntimeError(msg)
    if stderr:
        msg = f"Excel benchmark wrote an unexpected error stream for {workbook.name}: {stderr}"
        raise RuntimeError(msg)
    pairs = re.findall(r"([A-Z_]+)=([0-9]+(?:\.[0-9]+)?)", stdout)
    metrics = {name: float(value) for name, value in pairs}
    missing = EXPECTED_METRICS - set(metrics)
    extra = set(metrics) - EXPECTED_METRICS
    if missing or extra:
        msg = (
            f"Excel benchmark returned {stdout!r}; missing={sorted(missing)}, extra={sorted(extra)}"
        )
        raise RuntimeError(msg)
    return metrics


def populated_items(row_count: int) -> list[dict[str, object]]:
    """Return a representative populated hierarchy for formula-scale timing.

    Returns:
        Representative Item records at the requested scale.

    Raises:
        ValueError: If fewer than two rows are requested.

    """
    if row_count < MIN_SCALE_ROWS:
        msg = "the populated benchmark requires at least two Items rows"
        raise ValueError(msg)
    today = datetime.now(tz=UTC).astimezone().date()
    rows = [
        {
            "ID": "I-800001",
            "Title": "Performance portfolio",
            "Type": "Project",
            "Status": "In Progress",
            "Delivery Health": "On track",
            "Priority": "P1",
            "Owner": "Scale Owner",
            "Start": today - timedelta(days=30),
            "Due": today + timedelta(days=180),
            "Latest Status": "Representative populated performance model.",
            "Created": today - timedelta(days=60),
            "Updated": today,
            "InProgressSince": today - timedelta(days=30),
            "LatestUpdateOn": today,
        }
    ]
    statuses = ("In Progress", "Ready", "Done")
    health = ("On track", "At risk", "Off track")
    for offset in range(1, row_count):
        status = statuses[offset % len(statuses)]
        item = {
            "ID": f"I-{800001 + offset}",
            "Title": f"Performance work item {offset:04d}",
            "Type": "Task",
            "Parent": "I-800001",
            "Status": status,
            "Delivery Health": health[offset % len(health)],
            "Priority": f"P{offset % 5}",
            "Owner": "Scale Owner",
            "Start": today - timedelta(days=offset % 45),
            "Due": today + timedelta(days=(offset % 180) + 1),
            "Latest Status": "Representative populated performance model.",
            "Created": today - timedelta(days=90),
            "Updated": today - timedelta(days=offset % 7),
            "LatestUpdateOn": today - timedelta(days=offset % 7),
        }
        if status == "In Progress":
            item["InProgressSince"] = today - timedelta(days=(offset % 30) + 1)
        elif status == "Done":
            item["DoneDate"] = today - timedelta(days=(offset % 30) + 1)
        rows.append(item)
    return rows


def measure_populated_scale() -> dict[str, float]:
    """Build, calculate and measure the representative populated workbook.

    Returns:
        Parsed benchmark metrics for the populated workbook.

    """
    with temporary_examples() as examples, temporary_workbook("PM_performance_scale") as out:
        examples.ITEMS_EXAMPLES = populated_items(SCALE_ROWS)
        examples.PEOPLE_EXAMPLES = [
            {"Person": "Scale Owner", "Role": "Programme lead", "Team": "Core"},
        ]
        examples.RAID_EXAMPLES = []
        build_one(out, with_vba=False)
        recalculate(out)
        return measure(out, timeout=300)


def metric_failures(
    label: str,
    metrics: dict[str, float],
    limits: dict[str, float],
) -> list[str]:
    """Return every exceeded latency or state-restoration contract.

    Returns:
        Human-readable contract failures.

    """
    failures = [
        f"{label} {name}={metrics[name]:.1f}ms exceeds {limit:.1f}ms"
        for name, limit in limits.items()
        if metrics[name] > limit
    ]
    if metrics["OPEN_RETRIES"] > OPEN_RETRY_LIMIT:
        failures.append(
            f"{label} OPEN_RETRIES={metrics['OPEN_RETRIES']:.0f} exceeds {OPEN_RETRY_LIMIT}"
        )
    if metrics["CALCULATION_RESTORED"] != 1:
        failures.append(f"{label} did not restore Excel calculation mode")
    return failures


def print_metrics(label: str, metrics: dict[str, float], limits: dict[str, float]) -> None:
    """Print one compact benchmark result."""
    fields = [f"{name}={metrics[name]:.1f}ms/{limit:.1f}ms" for name, limit in limits.items()]
    fields.extend((
        f"OPEN_RETRIES={metrics['OPEN_RETRIES']:.0f}/{OPEN_RETRY_LIMIT}",
        f"CALCULATION_RESTORED={metrics['CALCULATION_RESTORED']:.0f}",
    ))
    sys.stdout.write(f"{label}: {' '.join(fields)}\n")


def main() -> None:
    """Run normal and populated performance gates for one workbook.

    Raises:
        SystemExit: If usage, script contracts, or latency limits fail.

    """
    if len(sys.argv) != EXPECTED_ARG_COUNT:
        msg = "Usage: python -m build.qa.performance workbook"
        raise SystemExit(msg)
    contract_failures = benchmark_contract_failures()
    if contract_failures:
        raise SystemExit("PERFORMANCE QA FAIL: " + "; ".join(contract_failures))
    path = Path(sys.argv[1])
    metrics = measure(path)
    scale_metrics = measure_populated_scale()
    failures = metric_failures(path.name, metrics, LIMITS_MS)
    failures.extend(metric_failures(f"populated-{SCALE_ROWS}", scale_metrics, SCALE_LIMITS_MS))
    print_metrics(path.name, metrics, LIMITS_MS)
    print_metrics(f"populated-{SCALE_ROWS}", scale_metrics, SCALE_LIMITS_MS)
    if failures:
        raise SystemExit("PERFORMANCE QA FAIL: " + "; ".join(failures))
    sys.stdout.write("PERFORMANCE QA PASS\n")


if __name__ == "__main__":
    main()
