"""Reserve dynamic-array spill zones and reject build-time overlaps.

Each formula reserves its worst-case rectangle so generated workbooks cannot
ship with predictable ``#SPILL!`` collisions.
"""

from dataclasses import dataclass, field


@dataclass(frozen=True, slots=True)
class SpillZone:
    """A rectangular spill reservation and its diagnostic label."""

    first_row: int
    first_column: int
    last_row: int
    last_column: int
    label: str

    def overlaps(self, other: "SpillZone") -> bool:
        """Return whether this reservation intersects another reservation.

        Returns:
            ``True`` when the two rectangles share at least one cell.

        """
        return not (
            self.last_row < other.first_row
            or other.last_row < self.first_row
            or self.last_column < other.first_column
            or other.last_column < self.first_column
        )


class _SpillOverlapError(ValueError):
    def __init__(self, sheet: str, zone: SpillZone, other: SpillZone) -> None:
        super().__init__(
            f"spill overlap on {sheet}: {zone.label} "
            f"({zone.first_row},{zone.first_column})-({zone.last_row},{zone.last_column}) "
            f"vs {other.label} "
            f"({other.first_row},{other.first_column})-({other.last_row},{other.last_column})"
        )


def _assert_available(sheet: str, zone: SpillZone, reservations: list[SpillZone]) -> None:
    """Reject a spill zone that intersects an existing reservation.

    Raises:
        _SpillOverlapError: If the proposed zone intersects an existing zone.

    """
    for other in reservations:
        if zone.overlaps(other):
            raise _SpillOverlapError(sheet, zone, other)


@dataclass(slots=True)
class SpillRegistry:
    """Track every spill reservation by worksheet name."""

    zones: dict[str, list[SpillZone]] = field(default_factory=dict)

    def reserve(self, sheet: str, zone: SpillZone) -> None:
        """Reserve a non-overlapping zone on a worksheet."""
        reservations = self.zones.setdefault(sheet, [])
        _assert_available(sheet, zone, reservations)
        reservations.append(zone)


REGISTRY = SpillRegistry()
