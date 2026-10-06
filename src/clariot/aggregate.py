"""Turns a pile of stored events into what a report needs to say.

Pure functions over ``StoredEvent`` lists: no database, no Outlook, no templates.

Events are stored with the report's original wording and translated here, at
report time. That way improving the glossary later also improves the wording of
events that were ingested before the phrase was known.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Sequence

from .models import AlertReport
from .store import StoredEvent
from .textutils import normalize


@dataclass(frozen=True)
class TimelineEntry:
    """One alarm as it appears in a report's chronology, already translated."""

    event_at: datetime | None
    event_type: str
    urgency: str
    equipment_status: str = ""
    possible_cause: str = ""
    recommended_action: str = ""


@dataclass(frozen=True)
class AlarmDetail:
    """One distinct condition on a machine, and every alarm that reported it.

    Grouped rather than listed one by one for a reason: a pump that alarms three
    times for the same looseness sends three reports with identical wording, and
    printing that paragraph three times reads like a mistake. What the client
    needs is the condition once and the dates it happened. When the conditions
    genuinely differ, each one gets its own block — which is the case the old
    "only the latest alarm" layout silently dropped.
    """

    event_type: str
    equipment_status: str
    possible_cause: str
    recommended_action: str
    occurrences: tuple[datetime | None, ...]
    """Newest first, same order as the timeline it came from."""

    @property
    def count(self) -> int:
        return len(self.occurrences)


@dataclass(frozen=True)
class EquipmentSummary:
    """Everything one machine did during the period."""

    equipment_key: str
    company: str
    plant: str
    machine: str
    serial_number: str
    machine_type: str
    events: tuple[StoredEvent, ...]
    """Newest first."""
    timeline: tuple[TimelineEntry, ...]
    """Same order as ``events``, with the wording already translated."""
    event_types: tuple[tuple[str, int], ...]
    """Distinct event types with their counts, most frequent first."""
    latest_status: str
    latest_cause: str
    latest_action: str
    urgency: str
    is_critical_urgency: bool

    @property
    def count(self) -> int:
        return len(self.events)

    @property
    def repeated(self) -> bool:
        return self.count > 1

    @property
    def first_at(self) -> datetime | None:
        stamps = [e.event_at for e in self.events if e.event_at]
        return min(stamps) if stamps else None

    @property
    def last_at(self) -> datetime | None:
        stamps = [e.event_at for e in self.events if e.event_at]
        return max(stamps) if stamps else None

    @property
    def event_keys(self) -> tuple[str, ...]:
        return tuple(e.event_key for e in self.events)

    @property
    def label(self) -> str:
        """How the machine is named in a report line."""
        name = self.machine or "Equipo sin identificar"
        return f"{name} ({self.serial_number})" if self.serial_number else name


def _translate(text: str, glossary) -> str:
    if not text or glossary is None:
        return text
    result = glossary.apply(AlertReport(fields={"equipment_status": text}))
    return result.report.fields["equipment_status"]


def _by_recency(events: Sequence[StoredEvent]) -> list[StoredEvent]:
    """Newest first. Events with no timestamp fall back to ingestion order."""
    return sorted(
        events,
        key=lambda e: (e.event_at.isoformat() if e.event_at else "", e.ingested_at),
        reverse=True,
    )


def distinct_details(entries: Sequence[TimelineEntry]) -> tuple[AlarmDetail, ...]:
    """Collapse a chronology into one block per distinct condition.

    Order follows first appearance in ``entries``, so passing a newest-first
    timeline puts the most recent condition first — which is what the client
    should read before the history.
    """
    groups: dict[tuple[str, str, str, str], list[datetime | None]] = {}
    for entry in entries:
        key = (
            entry.event_type,
            entry.equipment_status,
            entry.possible_cause,
            entry.recommended_action,
        )
        groups.setdefault(key, []).append(entry.event_at)

    return tuple(
        AlarmDetail(
            event_type=event_type,
            equipment_status=status,
            possible_cause=cause,
            recommended_action=action,
            occurrences=tuple(dates),
        )
        for (event_type, status, cause, action), dates in groups.items()
    )


def summarize_equipment(
    events: Sequence[StoredEvent],
    glossary=None,
    critical_urgencies: frozenset[str] = frozenset(),
) -> EquipmentSummary:
    """Collapse one machine's events into a single report block."""
    if not events:
        raise ValueError("summarize_equipment needs at least one event")

    ordered = _by_recency(events)
    latest = ordered[0]

    # Translate once per event and reuse: the same wording feeds both the
    # chronology and the type counts.
    timeline = tuple(
        TimelineEntry(
            event_at=event.event_at,
            event_type=_translate(event.event_type, glossary) or "Sin especificar",
            urgency=_translate(event.urgency, glossary),
            equipment_status=_translate(event.equipment_status, glossary),
            possible_cause=_translate(event.possible_cause, glossary),
            recommended_action=_translate(event.recommended_action, glossary),
        )
        for event in ordered
    )

    counts: dict[str, int] = {}
    for entry in timeline:
        counts[entry.event_type] = counts.get(entry.event_type, 0) + 1
    # Most frequent first; ties keep the order they were seen in.
    event_types = tuple(
        sorted(counts.items(), key=lambda item: item[1], reverse=True)
    )

    # The worst urgency wins, otherwise the most recent one. Arbitrary phrases
    # cannot be ranked, so only the configured critical values are treated as
    # more severe than "whatever the latest report said".
    critical = next(
        (e for e in ordered if e.urgency and normalize(e.urgency) in critical_urgencies),
        None,
    )
    source = critical or latest

    return EquipmentSummary(
        equipment_key=latest.equipment_key,
        company=latest.company,
        plant=latest.plant,
        machine=latest.machine,
        serial_number=latest.serial_number,
        machine_type=_translate(latest.machine_type, glossary),
        events=tuple(ordered),
        timeline=timeline,
        event_types=event_types,
        latest_status=_translate(latest.equipment_status, glossary),
        latest_cause=_translate(latest.possible_cause, glossary),
        latest_action=_translate(latest.recommended_action, glossary),
        urgency=_translate(source.urgency, glossary),
        is_critical_urgency=critical is not None,
    )


def group_by_equipment(
    events: Sequence[StoredEvent],
    glossary=None,
    critical_urgencies: frozenset[str] = frozenset(),
) -> list[EquipmentSummary]:
    """One summary per machine, ordered the way a technician should read them.

    Recurrence first, then declared urgency, then recency: a pump that alarmed
    four times outranks one that alarmed once, whatever the urgency field says.
    """
    buckets: dict[str, list[StoredEvent]] = {}
    for event in events:
        buckets.setdefault(event.equipment_key, []).append(event)

    summaries = [
        summarize_equipment(group, glossary, critical_urgencies)
        for group in buckets.values()
    ]
    return sorted(
        summaries,
        key=lambda s: (
            s.count,
            s.is_critical_urgency,
            s.last_at.isoformat() if s.last_at else "",
        ),
        reverse=True,
    )


def group_by_company(events: Sequence[StoredEvent]) -> dict[str, list[StoredEvent]]:
    """Split events per client, keyed by the company name as extracted.

    Grouping is case- and accent-insensitive so "NESTLÉ CHILE" and "Nestle Chile"
    land together, but the label kept is the first spelling seen, because that is
    what the report should print.
    """
    buckets: dict[str, list[StoredEvent]] = {}
    labels: dict[str, str] = {}
    for event in events:
        key = normalize(event.company) or "SIN-CLIENTE"
        labels.setdefault(key, event.company or "Cliente sin identificar")
        buckets.setdefault(key, []).append(event)
    return {labels[key]: group for key, group in buckets.items()}


def period_of(events: Sequence[StoredEvent]) -> tuple[datetime | None, datetime | None]:
    """Earliest and latest event timestamps in the set."""
    stamps = sorted(e.event_at for e in events if e.event_at)
    return (stamps[0], stamps[-1]) if stamps else (None, None)


MONTHS = (
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
)


def spanish_date(value: datetime | None, with_time: bool = False) -> str:
    """Dates in the report are written for a Chilean reader, not ISO."""
    if value is None:
        return "sin fecha"
    text = f"{value.day} de {MONTHS[value.month - 1]}"
    if with_time:
        text += f", {value:%H:%M}"
    return text


def spanish_period(start: datetime | None, end: datetime | None) -> str:
    """A date range as a Chilean reader would write it: "21 al 27 de julio"."""
    if start is None or end is None:
        return "periodo sin determinar"
    if start.date() == end.date():
        return spanish_date(start)
    if start.month == end.month:
        return f"{start.day} al {end.day} de {MONTHS[end.month - 1]}"
    return f"{spanish_date(start)} al {spanish_date(end)}"
