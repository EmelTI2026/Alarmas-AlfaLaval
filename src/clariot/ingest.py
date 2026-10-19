"""Capture stage: read new alerts, store the events, escalate what cannot wait.

Runs as often as convenient. Only alarms the classifier marks as immediate leave
the building today; everything else waits for the weekly report.
"""

from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from jinja2 import Environment

from .aggregate import summarize_equipment
from .audit import AuditLog
from .classifier import IMMEDIATE_LEVELS, SEVERITY, Classifier
from .config import ClientDirectory, Settings
from .models import AlertReport
from .ledger import STATUS_DONE, Ledger
from .pdf_parser import parse_pdf
from .report_builder import build_equipment_draft
from .store import EventStore, build_event, serial_from_subject
from .textutils import sanitize_filename

logger = logging.getLogger(__name__)


class _HeldMessage:
    """Stands in for the original email when retrying a held alarm.

    The email itself was already filed away; only its identity is still needed for
    the audit trail.
    """

    def __init__(self, event) -> None:
        self.message_key = event.message_key
        self.subject = f"(retenida) {event.machine_label}"
        self.entry_id = ""
        self.folder = ""


@dataclass
class IngestReport:
    """What one run did, for the log and the technician's console."""
    seen: int = 0
    stored: int = 0
    resent: int = 0
    skipped: int = 0
    ignored: int = 0
    updated_drafts: int = 0
    held: int = 0
    released: int = 0
    still_held: int = 0
    translation_failures: int = 0
    pending_recipient: int = 0
    needs_review: int = 0
    failed: int = 0
    errors: list[str] = field(default_factory=list)
    # Severity each machine's draft ENDED UP at, keyed by equipment. Counting
    # actions instead of drafts reported "4 normales" for a run that had in fact
    # left four [URGENTE] drafts: the first alarm of each machine was normal, and
    # the promotion on the second never corrected the tally. The console line is
    # what tells the technician which folder to open, so it has to describe the
    # drafts as they stand, not the order they got there.
    draft_levels: dict[str, str] = field(default_factory=dict)

    def _count(self, level: str) -> int:
        return sum(1 for value in self.draft_levels.values() if value == level)

    @property
    def critical_drafts(self) -> int:
        return self._count("critical")

    @property
    def urgent_drafts(self) -> int:
        return self._count("urgent")

    @property
    def single_drafts(self) -> int:
        return self._count("normal")

    def summary(self) -> str:
        parts = [
            f"{self.seen} en la carpeta",
            f"{self.stored} eventos nuevos",
        ]
        if self.resent:
            parts.append(f"{self.resent} reenvios (no contados)")
        if self.critical_drafts:
            parts.append(f"{self.critical_drafts} CRITICOS")
        if self.urgent_drafts:
            parts.append(f"{self.urgent_drafts} urgentes")
        if self.single_drafts:
            parts.append(f"{self.single_drafts} normales")
        if self.updated_drafts:
            parts.append(f"{self.updated_drafts} acoplados a un borrador abierto")
        if self.released:
            parts.append(f"{self.released} liberados del glosario")
        if self.held:
            parts.append(
                f"ATENCION: {self.held} RETENIDOS por frases sin traducir "
                "(ver --pending)"
            )
        if self.still_held:
            parts.append(f"{self.still_held} siguen retenidos")
        if self.translation_failures:
            parts.append(
                f"ATENCION: {self.translation_failures} PDF(s) sin traducir por falla "
                "del traductor (ver logs)"
            )
        if self.needs_review:
            parts.append(
                f"ATENCION: {self.needs_review} correo(s) con mas de un PDF, "
                "SIN PROCESAR (revisar a mano)"
            )
        if self.pending_recipient:
            parts.append(f"{self.pending_recipient} esperan destinatario")
        if self.skipped:
            parts.append(f"{self.skipped} omitidos")
        if self.ignored:
            parts.append(f"{self.ignored} sin PDF (no eran alertas)")
        if self.failed:
            parts.append(f"{self.failed} con error")
        return " | ".join(parts)


class Ingestor:
    """Reads new alerts and turns them into drafts.

    Collaborators are injected so the whole flow runs against fakes: nothing here
    imports Outlook or a translation service directly.
    """
    def __init__(
        self,
        *,
        settings: Settings,
        labels: dict,
        value_noise: tuple[str, ...],
        glossary,
        directory: ClientDirectory,
        outlook,
        store: EventStore,
        translator=None,
        ledger: Ledger,
        audit: AuditLog,
        env: Environment,
        classifier: Classifier | None = None,
        dry_run: bool = False,
    ) -> None:
        self.settings = settings
        self.labels = labels
        self.value_noise = value_noise
        self.glossary = glossary
        self.directory = directory
        self.outlook = outlook
        self.store = store
        self.translator = translator
        self.ledger = ledger
        self.audit = audit
        self.env = env
        self.dry_run = dry_run
        self.classifier = classifier or Classifier(
            window_days=settings.grouping.window_days,
            urgent_threshold=settings.grouping.urgent_threshold,
            same_day_is_critical=settings.grouping.same_day_is_critical,
            urgency_cooldown_days=settings.grouping.urgency_cooldown_days,
            critical_urgencies=settings.critical_urgencies,
        )

    # --------------------------------------------------------------------- run

    def run(self, limit: int | None = None) -> IngestReport:
        report = IngestReport()
        self._release_held(report)

        for source in self.settings.outlook.source_folders:
            messages = self.outlook.messages(
                source, self.settings.outlook.include_subfolders
            )
            logger.info("Found %s message(s) under '%s'", len(messages), source)

            for message in messages:
                if limit is not None and report.seen >= limit:
                    logger.info("Reached the --limit of %s message(s)", limit)
                    return report
                report.seen += 1

                if not self.ledger.should_process(message.message_key):
                    report.skipped += 1
                    self._file_away(message, source, already_done=True)
                    continue

                try:
                    self._process(message, source, report)
                except Exception as exc:  # noqa: BLE001 - one bad alert must not stop the batch
                    logger.exception("Failed to ingest '%s'", message.subject)
                    report.failed += 1
                    report.errors.append(f"{message.subject}: {exc}")
                    if not self.dry_run:
                        self.ledger.mark_failed(message.message_key, str(exc))
                    self.audit.record(
                        message_key=message.message_key,
                        original_subject=message.subject,
                        status="error",
                        detail=str(exc)[:500],
                    )

        return report

    def _release_held(self, report: IngestReport) -> None:
        """Retry alarms held back earlier, now that the glossary may cover them.

        Runs before reading new mail, so adding a phrase to the glossary and
        double-clicking the shortcut is all it takes to unblock what was waiting.
        """
        held = self.store.blocked()
        if not held:
            return
        logger.info("%s alarm(s) were held back; retrying", len(held))

        for event in held:
            if self._glossary_gaps([event]):
                report.still_held += 1
                continue
            try:
                fake = _HeldMessage(event)
                self._classify_and_escalate(event, report, fake)
                report.released += 1
            except Exception:  # noqa: BLE001 - one bad alarm must not stop the rest
                logger.exception("Could not release the held alarm %s", event.event_key)
                report.still_held += 1

    # ---------------------------------------------------------------- one item

    def _process(self, message, source: str, report: IngestReport) -> None:
        logger.info("Ingesting: %s", message.subject)
        if not self.dry_run:
            self.ledger.mark_in_progress(
                message.message_key, message.entry_id, message.subject
            )

        work_dir = self.settings.paths.work_dir / sanitize_filename(
            message.message_key, fallback="mensaje"
        )
        shutil.rmtree(work_dir, ignore_errors=True)
        work_dir.mkdir(parents=True, exist_ok=True)

        try:
            pdfs = self.outlook.save_pdf_attachments(
                message.entry_id, work_dir, self.settings.outlook.max_attachment_mb
            )
            if not pdfs:
                # The Outlook rule filters by sender, so account notices from the
                # same provider land here too. Not an error.
                logger.warning("No PDF in '%s'; not an alert", message.subject)
                report.ignored += 1
                self.audit.record(
                    message_key=message.message_key,
                    original_subject=message.subject,
                    status="ignorado",
                    detail="el correo no trae PDF adjunto; no es una alerta",
                )
                self._complete(message, source)
                return

            if len(pdfs) > 1:
                # A real alert carries exactly one report. Two means something
                # changed at the source, and there is no honest way to pick: using
                # the first would silently drop an alarm, and the whole point of
                # this system is counting alarms. So it is not processed at all.
                #
                # Left where it is on purpose, with the ledger row still open, so
                # every run reports it again until a person looks at it. Filing it
                # away would make it disappear quietly, which is the failure mode
                # this branch exists to prevent.
                logger.error(
                    "'%s' trae %s PDFs adjuntos y no se procesara: %s. "
                    "Revisar a mano en la carpeta de origen.",
                    message.subject,
                    len(pdfs),
                    ", ".join(p.name for p in pdfs),
                )
                report.needs_review += 1
                report.errors.append(
                    f"{message.subject}: {len(pdfs)} PDFs adjuntos, sin procesar"
                )
                self.audit.record(
                    message_key=message.message_key,
                    original_subject=message.subject,
                    status="revisar",
                    detail=(
                        f"{len(pdfs)} PDFs adjuntos; un reporte real trae uno solo, "
                        "no se procesa para no perder una alarma"
                    ),
                )
                return

            alert = parse_pdf(pdfs[0], self.labels, self.value_noise)
            alert = self._serial_fallback(alert, message.subject)

            if self.dry_run:
                # A dry run writes nothing anywhere: no archive, no event, no
                # draft. Otherwise the rehearsal would make the real run treat the
                # alarm as an already-counted resend.
                event = build_event(
                    alert, message_key=message.message_key, pdf_path=str(pdfs[0])
                )
                report.stored += 1
                self._classify_and_escalate(event, report, message)
                return

            archived = self._archive(pdfs[0], alert)
            translated = self._translate_pdf(archived, report)
            event, is_new = self.store.record(
                alert,
                message_key=message.message_key,
                pdf_path=str(archived),
                pdf_translated=str(translated) if translated else "",
            )

            if not is_new:
                # A resent notification about an event already counted. The PDF is
                # archived anyway; the alarm is not counted twice.
                report.resent += 1
                self.audit.record(
                    message_key=message.message_key,
                    original_subject=message.subject,
                    company=event.company,
                    machine=event.machine,
                    status="reenvio",
                    detail=f"evento ya registrado el {event.event_date_raw}",
                )
                self._complete(message, source)
                return

            report.stored += 1
            self._classify_and_escalate(event, report, message)
            self._complete(message, source)
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    # -------------------------------------------------------------- escalation

    def _classify_and_escalate(self, event, report: IngestReport, message) -> None:
        # The full history of the machine, with no date pre-filter. The classifier
        # windows relative to the *event's* own timestamp, not to now, and
        # pre-filtering by "now minus the window" silently dropped that history for
        # any alarm that arrives late or sat unprocessed for a few days — so three
        # alarms on one pump looked like three isolated ones. Retention keeps this
        # list short.
        history = self.store.history_for(event.equipment_key)
        verdict = self.classifier.classify(
            event,
            history,
            last_urgency_report=self.store.last_urgency_report(event.equipment_key),
            last_urgency_level=self.store.last_urgency_level(event.equipment_key),
        )
        logger.info("Level %s for %s: %s", verdict.level, event.machine, verdict.reason)

        # A client-facing report must never carry untranslated text. If the
        # glossary does not cover this alarm, no draft is built: the alarm waits
        # and the next run picks it up once the phrase is added. Safety alarms are
        # the exception — delaying one over a dictionary gap would be worse.
        missing = self._glossary_gaps(verdict.events)
        holdable = not (verdict.immediate and self.settings.glossary.never_hold_urgent)
        if missing and self.settings.glossary.holds and holdable:
            phrases = "; ".join(f"{field}={text!r}" for field, text in missing)
            logger.warning(
                "Draft for %s held back: %s phrase(s) missing from the glossary. %s",
                event.machine,
                len(missing),
                phrases,
            )
            self.store.mark_blocked(
                event.event_key, f"{len(missing)} frase(s) sin traducir: {phrases}"
            )
            report.held += 1
            self.audit.record(
                message_key=message.message_key,
                original_subject=message.subject,
                company=event.company,
                machine=event.machine,
                urgency=event.urgency,
                status="retenido",
                detail=f"faltan {len(missing)} frase(s) en el glosario: {phrases}"[:500],
            )
            return

        if not verdict.immediate and not self.settings.reports.single_alarm_as_draft:
            # Accumulate for the weekly report.
            self.audit.record(
                message_key=message.message_key,
                original_subject=message.subject,
                company=event.company,
                machine=event.machine,
                urgency=event.urgency,
                status="acumulado",
                detail=f"{verdict.reason}; sale en el informe semanal",
            )
            return

        # An open draft for this machine absorbs the new alarm instead of
        # producing a second email. That is the whole point: the client gets one
        # message per machine per week, and it grows as alarms arrive.
        existing = self._reusable_draft(event)

        # A draft's severity is the worst the machine reached, never the level of
        # whichever alarm happened to arrive last. The cooldown downgrades a repeat
        # to "normal" to avoid mailing the client twice — but applying that to a
        # draft being *rewritten* renamed an [URGENTE] back to a routine report and
        # moved it out of the urgency folder. Three alarms on one pump looked like
        # one.
        level = verdict.level
        if existing:
            level = self._worst_level(level, self.store.open_draft_level(event.equipment_key))

        summary = summarize_equipment(
            verdict.events, self.glossary, self.settings.critical_urgencies
        )
        draft = build_equipment_draft(
            summary,
            level,
            self.directory,
            self.settings,
            self.env,
            include_reported=existing is not None,
            missing_phrases=missing,
        )

        if self.dry_run:
            action = "update" if existing else "create"
            logger.info("[dry-run] Would %s a %s draft: %s", action, level, draft.subject)
            return

        urgent = level in IMMEDIATE_LEVELS
        folder = (
            self.settings.outlook.urgent_draft_folder
            or self.settings.outlook.draft_folder
            if urgent
            else self.settings.outlook.draft_folder
        )

        if existing:
            entry_id = self.outlook.update_draft(
                existing,
                subject=draft.subject,
                html_body=draft.html_body,
                cc=draft.cc,
                attachments=draft.attachments,
                target_folder=folder,
                target_folder_parent=self.settings.outlook.draft_folder_parent,
            )
            self.store.remember_draft(
                event.equipment_key, entry_id, draft.kind, summary.first_at
            )
            self.store.mark_reported(draft.event_keys, draft.kind)
            report.updated_drafts += 1
            report.draft_levels[event.equipment_key] = level
            self.audit.record(
                message_key=message.message_key,
                original_subject=message.subject,
                company=event.company,
                machine=event.machine,
                urgency=event.urgency,
                draft_subject=draft.subject,
                status=f"acoplado_{draft.kind}",
                detail=f"{verdict.reason}; se sumo al borrador abierto",
            )
            return

        entry_id = self.outlook.create_draft(
            subject=draft.subject,
            html_body=draft.html_body,
            to=draft.to,
            cc=draft.cc,
            attachments=draft.attachments,
            sender_account=self.settings.email.sender_account,
            target_folder=folder,
            target_folder_parent=self.settings.outlook.draft_folder_parent,
        )
        # Marked reported the instant the draft exists, so a failure below can
        # never produce a second one, and so these events do not appear again in
        # any later report.
        self.store.mark_reported(draft.event_keys, draft.kind)
        self.store.remember_draft(
            event.equipment_key, entry_id, draft.kind, summary.first_at
        )

        report.draft_levels[event.equipment_key] = level
        if not draft.client_resolved:
            report.pending_recipient += 1

        self.audit.record(
            message_key=message.message_key,
            original_subject=message.subject,
            company=event.company,
            machine=event.machine,
            urgency=event.urgency,
            recipients="; ".join(draft.to),
            draft_subject=draft.subject,
            status=f"informe_{draft.kind}",
            detail=verdict.reason,
        )

    def _serial_fallback(self, alert: AlertReport, subject: str) -> AlertReport:
        """Take the serial from the subject when the PDF did not yield one.

        Grouping keys on the serial. If the report layout ever changes and the
        parser stops finding it, repeated alarms on one pump would stop being
        detected — silently. The subject carries the same value, so it is a free
        second source.
        """
        if alert.fields.get("serial_number"):
            return alert

        serial = serial_from_subject(subject)
        if not serial:
            return alert

        logger.info("Serial taken from the subject: %r", serial)
        return AlertReport(fields={**alert.fields, "serial_number": serial})

    def _glossary_gaps(self, events) -> list[tuple[str, str]]:
        """Phrases of these events that the glossary does not cover."""
        if self.glossary is None:
            return []
        gaps: dict[tuple[str, str], None] = {}
        for event in events:
            for field, text in self.glossary.apply(
                AlertReport(
                    fields={
                        "event_type": event.event_type,
                        "equipment_status": event.equipment_status,
                        "possible_cause": event.possible_cause,
                        "recommended_action": event.recommended_action,
                        "urgency": event.urgency,
                        "machine_type": event.machine_type,
                    }
                )
            ).missing:
                gaps[(field, text)] = None
        return list(gaps)

    @staticmethod
    def _worst_level(nuevo: str, anterior: str | None) -> str:
        """The higher of two severities. A draft never gets downgraded."""
        if not anterior:
            return nuevo
        return nuevo if SEVERITY.get(nuevo, 0) >= SEVERITY.get(anterior, 0) else anterior

    def _reusable_draft(self, event) -> str | None:
        """EntryID of the open draft this alarm should join, if any.

        Four reasons to start a new draft instead:
        the week is over, the draft was sent, it was deleted, or the technician
        already typed into it — rewriting it would destroy his work.
        """
        record = self.store.open_draft(event.equipment_key)
        if record is None:
            return None
        entry_id, _, first_at = record

        reference = event.event_at or datetime.now()
        if first_at is not None:
            if abs((reference - first_at).days) > self.settings.grouping.window_days:
                logger.info("The open draft for %s is older than the window", event.machine)
                self.store.forget_draft(event.equipment_key)
                return None

        route = self.directory.resolve(getattr(event, "company", ""))
        state = self.outlook.draft_state(
            entry_id,
            expected_cc=self._expected_cc(event),
            expected_to=route.to if route is not None else (),
        )
        if state == self.outlook.DRAFT_CLEAN:
            return entry_id

        logger.info(
            "Not joining the open draft for %s: %s. A new draft will be created.",
            event.machine,
            {
                self.outlook.DRAFT_SENT: "ya fue enviado",
                self.outlook.DRAFT_TOUCHED: "el tecnico ya lo edito",
                self.outlook.DRAFT_MISSING: "ya no existe",
            }.get(state, state),
        )
        self.store.forget_draft(event.equipment_key)
        return None

    def _expected_cc(self, event) -> tuple[str, ...]:
        """Every address the program itself could have copied on this machine.

        The union, not the exact list: the severity is decided further down, after
        the draft has already been chosen for reuse, so the critical copies are
        counted in even when this alarm turns out to be routine. Being permissive
        here can at most miss a manual CC edit; being strict would make our own
        copies look like one and break the "one growing draft per equipment" rule.
        """
        addresses: list[str] = []
        route = self.directory.resolve(getattr(event, "company", ""))
        if route is not None:
            addresses.extend(route.cc)
        addresses.extend(self.directory.always_cc)
        addresses.extend(self.directory.critical_cc)

        seen: set[str] = set()
        unique: list[str] = []
        for address in addresses:
            key = address.lower()
            if key not in seen:
                seen.add(key)
                unique.append(address)
        return tuple(unique)

    # ------------------------------------------------------------- housekeeping

    def _translate_pdf(self, original: Path, report: IngestReport) -> Path | None:
        """Translate the archived report. Returns the translated file, or None.

        A translation outage must not cost the alarm: the draft is built anyway
        with the original attached, and the summary says so in capitals.
        """
        if self.translator is None or self.dry_run:
            return None

        destination = original.with_name(f"{original.stem}_ES.pdf")
        if destination.exists():
            return destination
        try:
            self.translator.translate(original, destination)
            return destination
        except Exception as exc:  # noqa: BLE001
            report.translation_failures += 1
            logger.error(
                "No se pudo traducir %s; se adjuntara el original. Causa: %s",
                original.name,
                exc,
            )
            return None

    def _archive(self, pdf: Path, alert) -> Path:
        """Keep the report exactly as it arrived. This copy is what gets attached.

        The original filename is preserved on purpose. Alfa Laval already names
        the file with everything worth knowing —
        ``2026_07_24_Soprole_Soprole - San Bernardo Planta 4 VX_VX-3037575.pdf``,
        that is date, company, plant and serial — so renaming it would throw
        information away and make the attachment the client receives look like
        something we generated rather than the vendor's own report.
        """
        month_dir = self.settings.paths.archive_dir / datetime.now().strftime("%Y-%m")
        month_dir.mkdir(parents=True, exist_ok=True)

        base = month_dir / sanitize_filename(pdf.name, fallback="reporte.pdf")
        incoming = pdf.read_bytes()
        destination = base
        suffix = 1

        # Two different reports can share a filename. Each candidate is checked
        # until a free slot is found, never overwritten: a timestamp suffix alone
        # would collide when several alarms are processed inside the same second,
        # and one alarm's draft would end up carrying another's PDF.
        while destination.exists():
            if destination.read_bytes() == incoming:
                # The same report archived twice — a resent notification. Reuse it
                # instead of piling up identical copies.
                return destination
            destination = base.with_name(f"{base.stem}_{suffix}{base.suffix}")
            suffix += 1

        shutil.copy2(pdf, destination)
        return destination

    def _processed_target(self, message, source: str) -> str:
        base = self.settings.outlook.processed_folder
        if not base:
            return ""
        origin = getattr(message, "folder", "") or ""
        if origin.lower().startswith(source.lower()):
            relative = origin[len(source) :].strip("/")
            if relative:
                return f"{base}/{relative}"
        return base

    def _complete(self, message, source: str) -> None:
        if self.dry_run:
            return
        self.ledger.mark_done(message.message_key)
        self._file_away(message, source)

    def _file_away(self, message, source: str, already_done: bool = False) -> None:
        if self.dry_run:
            return
        if already_done and self.ledger.status(message.message_key) != STATUS_DONE:
            # Failed past its retry budget: leave it visible for a human.
            return
        try:
            self.outlook.finish_message(
                message.entry_id,
                mark_read=self.settings.outlook.mark_as_read,
                move_to=self._processed_target(message, source),
            )
        except Exception:  # noqa: BLE001
            logger.warning("Could not file '%s' away", message.subject)
