"""Builds the two report drafts. Knows nothing about Outlook or the database."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from jinja2 import Environment, FileSystemLoader, select_autoescape

from .aggregate import (
    MONTHS,
    EquipmentSummary,
    distinct_details,
    period_of,
    spanish_date,
    spanish_period,
)
from .classifier import LEVEL_CRITICAL, LEVEL_URGENT
from .config import ClientDirectory, Settings
from .models import ClientRoute
from .store import REPORT_CRITICAL, REPORT_SINGLE, REPORT_URGENT, REPORT_WEEKLY

logger = logging.getLogger(__name__)

URGENCY_TEMPLATE = "urgency_report.html.j2"
WEEKLY_TEMPLATE = "weekly_report.html.j2"


def build_template_env(templates_dir: Path) -> Environment:
    """Jinja environment for the report templates, with autoescaping on."""
    return Environment(
        loader=FileSystemLoader(str(templates_dir)),
        autoescape=select_autoescape(["html", "xml", "j2"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )


@dataclass(frozen=True)
class ReportDraft:
    """Everything needed to create or update the draft, with no Outlook involved."""
    subject: str
    html_body: str
    to: tuple[str, ...]
    cc: tuple[str, ...]
    attachments: tuple[Path, ...]
    client_resolved: bool
    event_keys: tuple[str, ...]
    kind: str
    company: str

    @property
    def critical(self) -> bool:
        return self.kind == REPORT_CRITICAL


def _when(occurrences: Sequence) -> str:
    """When a condition happened, written for a reader and not for a machine.

    Chronological even though the occurrences arrive newest first: a list of dates
    reads backwards otherwise. Four or more collapse to a range, because by then
    the enumeration is longer than it is useful and the timeline table above
    already has every timestamp.
    """
    dates = sorted(d for d in occurrences if d is not None)
    missing = len(occurrences) - len(dates)

    if not dates:
        return "sin fecha"
    if len(dates) == 1:
        text = spanish_date(dates[0], with_time=True)
    elif len(dates) > 3:
        text = spanish_period(dates[0], dates[-1])
    elif len({(d.year, d.month) for d in dates}) == 1:
        # Same month: "22, 23 y 24 de julio", not "22 de julio y 23 de julio".
        days = [str(d.day) for d in dates]
        text = f"{', '.join(days[:-1])} y {days[-1]} de {MONTHS[dates[0].month - 1]}"
    else:
        written = [spanish_date(d) for d in dates]
        text = f"{', '.join(written[:-1])} y {written[-1]}"

    if missing:
        text += f" (+{missing} sin fecha)"
    return text


def _recipients(
    route: ClientRoute | None, directory: ClientDirectory, critical: bool
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    to = route.to if route else ()
    cc = list(route.cc) if route else []
    cc.extend(directory.always_cc)
    if critical:
        cc.extend(directory.critical_cc)

    seen = {address.lower() for address in to}
    unique: list[str] = []
    for address in cc:
        key = address.lower()
        if key not in seen:
            seen.add(key)
            unique.append(address)
    return tuple(to), tuple(unique)


def _attachments(
    events: Sequence, limit: int = 20, attach_original: bool = False
) -> tuple[Path, ...]:
    """PDFs to attach, deduplicated and newest first.

    The translated report is what the client should read; the original goes along
    only when ``attach_original`` is on, because two copies of the same report is
    confusing and multiplies with every alarm in the draft.

    Capped: a client with fifty alarms must not get a fifty-attachment email that
    the mail server would bounce.
    """
    paths: list[Path] = []
    seen: set[str] = set()
    for event in events:
        candidates = []
        translated = getattr(event, "pdf_translated", "")
        if translated:
            candidates.append(translated)
        if event.pdf_path and (attach_original or not translated):
            candidates.append(event.pdf_path)

        for candidate in candidates:
            path = Path(candidate)
            if path.name in seen or not path.exists():
                continue
            seen.add(path.name)
            paths.append(path)

    if len(paths) > limit:
        logger.warning(
            "%s reports available, attaching the %s most recent",
            len(paths),
            limit,
        )
        paths = paths[:limit]
    return tuple(paths)


def _client_name(settings: Settings) -> str:
    """The greeting names the person who receives the mail, not the company.

    The technician knows the contact and types it in, so the placeholder stays
    even when the client is known and its recipients are pre-loaded.
    """
    return settings.email.greeting_placeholder


def _format_subject(template: str, **values) -> str:
    try:
        return template.format(**values).strip()
    except KeyError as exc:
        raise ValueError(
            f"El asunto configurado usa un campo desconocido {exc}. "
            "Disponibles: machine, count, period, date, company, plant"
        ) from exc


def build_equipment_draft(
    summary: EquipmentSummary,
    level: str,
    directory: ClientDirectory,
    settings: Settings,
    env: Environment,
    include_reported: bool = False,
    missing_phrases: Sequence[tuple[str, str]] = (),
) -> ReportDraft:
    """One machine and its alarms, as a draft ready for review.

    Covers all three levels. The level decides the subject, the tone of the
    opening paragraph and which folder the caller files the draft into.

    ``include_reported`` distinguishes the two situations:

    * ``True``  — an unsent draft is being rewritten, so it must carry **every**
      alarm of the window with all their PDFs together. Nothing reached the
      client yet, so nothing is a repeat.
    * ``False`` — a fresh draft after an earlier one was already sent. Only the
      new alarms are detailed; resending what the client already has reads like a
      mistake, and a context line explains where the severity comes from.
    """
    critical = level == LEVEL_CRITICAL
    urgent = level == LEVEL_URGENT
    route = directory.resolve(summary.company)
    to, cc = _recipients(route, directory, critical)
    start, end = summary.first_at, summary.last_at

    if include_reported:
        fresh = list(summary.timeline)
        fresh_events = list(summary.events)
    else:
        fresh = [
            entry
            for entry, event in zip(summary.timeline, summary.events)
            if not event.reported
        ]
        fresh_events = [event for event in summary.events if not event.reported]
    already_reported = summary.count - len(fresh_events)

    identification = [
        (label, value)
        for label, value in (
            ("Cliente", summary.company),
            ("Planta", summary.plant),
            ("Equipo", summary.machine),
            ("N° de serie", summary.serial_number),
            ("Tipo de equipo", summary.machine_type),
        )
        if value
    ]

    timeline = [
        {
            "date": spanish_date(entry.event_at, with_time=True),
            "event_type": entry.event_type,
            "urgency": entry.urgency,
        }
        for entry in fresh
    ]

    # One detail block per distinct condition, built from ``fresh`` and not from
    # the whole summary: after a draft was sent, the client must not be told again
    # about what he already read.
    details = [
        {
            "event_type": detail.event_type,
            "equipment_status": detail.equipment_status,
            "possible_cause": detail.possible_cause,
            "recommended_action": detail.recommended_action,
            "count": detail.count,
            "when": _when(detail.occurrences),
        }
        for detail in distinct_details(fresh)
    ]

    if critical:
        template = settings.reports.critical_subject
    elif urgent:
        template = settings.reports.urgent_subject
    else:
        template = settings.reports.normal_subject
    subject = _format_subject(
        template,
        machine=summary.machine or "Equipo sin identificar",
        count=summary.count,
        period=spanish_period(start, end),
        date=spanish_date(end),
        company=summary.company,
        plant=summary.plant,
    )
    if not route:
        subject = settings.email.unknown_client_subject_prefix + subject

    # A safety alarm is never held back over a dictionary gap, so its draft can
    # carry untranslated text. It must then be impossible to send by accident:
    # the subject is prefixed and a blocking banner sits at the top of the body.
    if missing_phrases:
        subject = f"[TRADUCIR] {subject}"

    attachments = _attachments(
        fresh_events, attach_original=settings.translation.attach_original
    )
    html = env.get_template(URGENCY_TEMPLATE).render(
        greeting=settings.email.greeting,
        client_name=_client_name(settings),
        summary=summary,
        critical=critical,
        urgent=urgent,
        period=spanish_period(start, end),
        identification=identification,
        timeline=timeline,
        details=details,
        new_count=len(fresh_events),
        already_reported=already_reported,
        first_date=spanish_date(start),
        emeltec_note=settings.reports.include_emeltec_note,
        brand=settings.reports.brand_color,
        missing_phrases=[text for _, text in missing_phrases],
    )

    return ReportDraft(
        subject=subject,
        html_body=html,
        to=to,
        cc=cc,
        attachments=attachments,
        # Only the new alarms are marked: the earlier ones keep the report they
        # actually went out in, which is what the audit trail should say.
        event_keys=tuple(event.event_key for event in fresh_events),
        client_resolved=route is not None,
        kind=REPORT_CRITICAL if critical else (REPORT_URGENT if urgent else REPORT_SINGLE),
        company=summary.company,
    )


def build_weekly_draft(
    company: str,
    summaries: Sequence[EquipmentSummary],
    directory: ClientDirectory,
    settings: Settings,
    env: Environment,
) -> ReportDraft:
    """Every isolated alarm of one client over the period, in one email."""
    if not summaries:
        raise ValueError("build_weekly_draft needs at least one equipment summary")

    route = directory.resolve(company)
    to, cc = _recipients(route, directory, critical=False)

    events = [event for summary in summaries for event in summary.events]
    start, end = period_of(events)

    def row(summary: EquipmentSummary) -> dict:
        return {
            "label": summary.label,
            "plant": summary.plant,
            "count": summary.count,
            "status": summary.latest_status or "Sin especificar",
            "types": ", ".join(
                f"{name} ({count})" if count > 1 else name
                for name, count in summary.event_types
            ),
            "urgency": summary.urgency or "",
            "date": spanish_date(summary.last_at),
        }

    # Repeated equipment first and in its own table: that is the finding the
    # client cannot see from the individual notifications.
    repeated = [row(s) for s in summaries if s.repeated]
    single = [row(s) for s in summaries if not s.repeated]

    subject = _format_subject(
        settings.reports.weekly_subject,
        company=company,
        period=spanish_period(start, end),
        count=len(events),
        machine="",
        date=spanish_date(end),
        plant=summaries[0].plant,
    )
    if not route:
        subject = settings.email.unknown_client_subject_prefix + subject

    attachments = _attachments(
        events, attach_original=settings.translation.attach_original
    )
    html = env.get_template(WEEKLY_TEMPLATE).render(
        greeting=settings.email.greeting,
        client_name=_client_name(settings),
        period=spanish_period(start, end),
        total_events=len(events),
        summaries=summaries,
        repeated=repeated,
        single=single,
        emeltec_note=settings.reports.include_emeltec_note,
        brand=settings.reports.brand_color,
    )

    return ReportDraft(
        subject=subject,
        html_body=html,
        to=to,
        cc=cc,
        attachments=attachments,
        client_resolved=route is not None,
        event_keys=tuple(event.event_key for event in events),
        kind=REPORT_WEEKLY,
        company=company,
    )
