"""The capture flow end to end, with fakes standing in for Outlook.

These are the guarantees the whole project rests on: an alarm already handled must
never produce a second draft, and an alarm never handled must never be lost.
"""

from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pytest

from clariot.audit import AuditLog
from clariot.glossary import Glossary
from clariot.ingest import Ingestor
from clariot.ledger import Ledger
from clariot.models import AlertReport
from clariot.report_builder import build_template_env
from clariot.store import EventStore

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def message(key="<alerta-1@aliot>", folder="1_Auditoria_Clariot"):
    return SimpleNamespace(
        entry_id="ENTRY-1",
        message_key=key,
        subject="Event notification report - VX-3037575",
        received="2026-07-24 22:34",
        sender="support@aliotportal.com",
        folder=folder,
    )


class FakeOutlook:
    """Records what was asked of Outlook, and can be told to fail on demand."""

    DRAFT_MISSING = "missing"
    DRAFT_SENT = "sent"
    DRAFT_TOUCHED = "touched"
    DRAFT_CLEAN = "clean"

    def __init__(self, messages, *, fail_on_finish=False, attachments=True, state=None):
        self._messages = messages
        self.fail_on_finish = fail_on_finish
        self.attachments = attachments
        self._state = state or self.DRAFT_CLEAN
        self.drafts = []
        self.updates = []
        self.finished = []
        self.state_queries = []
        self.to_queries = []

    def messages(self, folder, include_subfolders=True):
        return list(self._messages)

    def save_pdf_attachments(self, entry_id, dest_dir, max_mb=20):
        if not self.attachments:
            return []
        dest_dir.mkdir(parents=True, exist_ok=True)
        pdf = dest_dir / "reporte.pdf"
        pdf.write_bytes(b"%PDF-1.4 fake")
        return [pdf]

    def create_draft(self, **kwargs):
        self.drafts.append(kwargs)
        return f"DRAFT-{len(self.drafts)}"

    def draft_state(self, entry_id, *, expected_cc=(), expected_to=()):
        self.state_queries.append((entry_id, tuple(expected_cc)))
        self.to_queries.append((entry_id, tuple(expected_to)))
        return self._state

    def update_draft(self, entry_id, **kwargs):
        self.updates.append({"entry_id": entry_id, **kwargs})
        return entry_id

    def finish_message(self, entry_id, mark_read=True, move_to=""):
        if self.fail_on_finish:
            raise RuntimeError("Outlook lost the connection while filing")
        self.finished.append((entry_id, mark_read, move_to))


class FakeTranslator:
    def __init__(self, fail=False):
        self.fail = fail

    def translate(self, source, destination):
        if self.fail:
            raise RuntimeError("Billing account closed")
        destination.write_bytes(b"%PDF-1.4 traducido")
        return destination


ALERT = {
    "company": "Soprole",
    "plant": "Soprole - San Bernardo Planta 4",
    "machine": "VX-RetCIPLin1 1206230",
    "serial_number": "VX-3037575",
    "machine_type": "Centrifugal_pump",
    "event_date": "24-07-2026 22:34",
    "event_type": "Mechanical",
    "equipment_status": "Potential impeller damage.",
    "recommended_action": "Open pump and check impeller.",
    "urgency": "Perform check at next planned stop.",
}


@pytest.fixture
def build(settings, directory, monkeypatch):
    """A factory for Ingestors sharing one store, with PDF parsing stubbed."""
    import clariot.ingest as ingest_module

    glossary = Glossary(
        {
            "Mechanical": "Mecánico",
            "Centrifugal_pump": "Bomba centrífuga",
            "Potential impeller damage.": "Posible daño del impulsor.",
            "Open pump and check impeller.": "Abrir la bomba y revisar el impulsor.",
            "Perform check at next planned stop.": "Revisar en la próxima parada.",
        }
    )
    store = EventStore(settings.paths.events_db)

    def factory(outlook, translator=None, alert=None, dry_run=False):
        monkeypatch.setattr(
            ingest_module,
            "parse_pdf",
            lambda path, labels, noise=(): AlertReport(fields=dict(alert or ALERT)),
        )
        return Ingestor(
            settings=settings,
            labels={},
            value_noise=(),
            glossary=glossary,
            directory=directory,
            outlook=outlook,
            store=store,
            translator=translator,
            ledger=Ledger(settings.paths.ledger_db),
            audit=AuditLog(settings.paths.audit_csv),
            env=build_template_env(PROJECT_ROOT / "templates"),
            dry_run=dry_run,
        )

    factory.store = store
    return factory


# --- the happy path ---------------------------------------------------------


def test_a_first_alarm_produces_one_draft(build):
    outlook = FakeOutlook([message()])
    report = build(outlook).run()

    assert report.stored == 1
    assert report.single_drafts == 1
    assert len(outlook.drafts) == 1
    assert outlook.finished == [("ENTRY-1", True, "2_Auditoria_Procesado")]


def test_the_draft_carries_the_spanish_report(build):
    outlook = FakeOutlook([message()])
    build(outlook).run()

    body = outlook.drafts[0]["html_body"]
    assert "Posible daño del impulsor." in body
    assert "Mecánico" in body
    assert "[EN]" not in body


def test_two_different_alarms_are_both_processed(build):
    outlook = FakeOutlook([message("<a@aliot>"), message("<b@aliot>")])
    factory = build(outlook)
    report = factory.run()

    # Same equipment and same event timestamp: the second is a resend.
    assert report.stored == 1
    assert report.resent == 1


# --- no duplicates ----------------------------------------------------------


def test_the_same_email_twice_creates_nothing_the_second_time(build):
    """The pile-up case: the folder fills while the PC is off, then two runs."""
    build(FakeOutlook([message()])).run()

    again = FakeOutlook([message()])
    report = build(again).run()

    assert again.drafts == []
    assert report.skipped == 1
    # Filed away so it stops being re-evaluated on every run.
    assert again.finished == [("ENTRY-1", True, "2_Auditoria_Procesado")]


def test_a_failure_after_the_draft_does_not_duplicate_it(build):
    """Outlook dies after the draft exists, before the original is filed."""
    outlook = FakeOutlook([message()], fail_on_finish=True)
    build(outlook).run()

    assert len(outlook.drafts) == 1
    assert outlook.finished == []

    retry = FakeOutlook([message()])
    report = build(retry).run()

    assert retry.drafts == []
    assert report.skipped == 1


def test_a_dry_run_leaves_no_trace(build):
    outlook = FakeOutlook([message()])
    build(outlook, dry_run=True).run()

    assert outlook.drafts == []
    assert outlook.finished == []

    real = FakeOutlook([message()])
    assert build(real).run().single_drafts == 1


# --- mail that is not an alert ----------------------------------------------


def test_mail_without_a_pdf_is_ignored_not_failed(build):
    """The Outlook rule filters by sender, so other provider mail lands here too."""
    outlook = FakeOutlook([message()], attachments=False)
    report = build(outlook).run()

    assert report.ignored == 1
    assert report.failed == 0
    assert outlook.drafts == []
    assert outlook.finished == [("ENTRY-1", True, "2_Auditoria_Procesado")]
    assert "no eran alertas" in report.summary()


def test_a_message_with_two_pdfs_is_not_processed_at_all(build):
    """A real alert carries one report. Two means nobody can honestly pick.

    Using the first would silently drop an alarm, and counting alarms is the whole
    point. It stays in the source folder and is reported on every run until a
    person looks at it.
    """

    class OutlookConDosPdfs(FakeOutlook):
        def save_pdf_attachments(self, entry_id, dest_dir, max_mb=20):
            dest_dir.mkdir(parents=True, exist_ok=True)
            paths = []
            for nombre in ("reporte_a.pdf", "reporte_b.pdf"):
                pdf = dest_dir / nombre
                pdf.write_bytes(b"%PDF-1.4 " + nombre.encode())
                paths.append(pdf)
            return paths

    outlook = OutlookConDosPdfs([message()])
    report = build(outlook).run()

    assert report.needs_review == 1
    assert outlook.drafts == []
    # Not filed away: filing it would make it disappear quietly.
    assert outlook.finished == []
    assert "SIN PROCESAR" in report.summary()

    # And it keeps coming back, run after run, until someone deals with it.
    otra_vez = OutlookConDosPdfs([message()])
    assert build(otra_vez).run().needs_review == 1


def test_an_ignored_message_is_not_reprocessed(build):
    build(FakeOutlook([message()], attachments=False)).run()

    again = FakeOutlook([message()], attachments=False)
    assert build(again).run().skipped == 1


# --- the PDF translation is not load-bearing --------------------------------


def test_a_translation_failure_still_produces_a_draft(build):
    """Expired credit or a service outage must not cost the draft."""
    outlook = FakeOutlook([message()])
    report = build(outlook, FakeTranslator(fail=True)).run()

    assert report.single_drafts == 1
    assert report.translation_failures == 1
    assert report.failed == 0
    # The archived original, not the temp copy: that is what gets attached.
    attached = [p.name for p in outlook.drafts[0]["attachments"]]
    assert len(attached) == 1 and attached[0].endswith(".pdf")
    assert "ATENCION" in report.summary()


# --- filing mirrors the source tree ----------------------------------------


def test_a_subfolder_is_mirrored_under_processed(build):
    outlook = FakeOutlook([message(folder="1_Auditoria_Clariot/Soprole")])
    build(outlook).run()

    assert outlook.finished == [("ENTRY-1", True, "2_Auditoria_Procesado/Soprole")]


def test_a_message_in_the_root_folder_goes_to_the_processed_root(build):
    outlook = FakeOutlook([message()])
    build(outlook).run()

    assert outlook.finished[0][2] == "2_Auditoria_Procesado"


def test_nested_subfolders_are_mirrored_in_full(build):
    outlook = FakeOutlook([message(folder="1_Auditoria_Clariot/Soprole/Planta4")])
    build(outlook).run()

    assert outlook.finished[0][2] == "2_Auditoria_Procesado/Soprole/Planta4"


def test_the_serial_falls_back_to_the_subject(build):
    """If the PDF layout changes and the serial is lost, grouping must survive."""
    sin_serie = {k: v for k, v in ALERT.items() if k != "serial_number"}
    outlook = FakeOutlook([message()])
    factory = build(outlook, alert=sin_serie)
    factory.run()

    stored = factory.store.pending() or factory.store._query("")
    assert stored[0].serial_number == "VX-3037575"
    assert stored[0].equipment_key.endswith("VX-3037575")


# --- the attached report must be the vendor's own file, untouched ------------


def test_the_archived_pdf_keeps_its_original_name(build):
    """The vendor's filename carries date, company, plant and serial."""
    outlook = FakeOutlook([message()])
    build(outlook).run()

    attached = outlook.drafts[0]["attachments"][0]
    assert attached.name == "reporte.pdf"


def test_the_archived_pdf_is_byte_identical(build):
    outlook = FakeOutlook([message()])
    build(outlook).run()

    attached = outlook.drafts[0]["attachments"][0]
    assert attached.read_bytes() == b"%PDF-1.4 fake"


def test_the_same_report_twice_is_not_archived_twice(build, settings):
    """A resent notification must not pile up identical copies."""
    build(FakeOutlook([message("<a@aliot>")])).run()
    build(FakeOutlook([message("<b@aliot>")])).run()

    archived = list(settings.paths.archive_dir.rglob("*.pdf"))
    assert len(archived) == 1


def test_an_alarm_already_opened_by_a_human_is_still_processed(build):
    """A technician who opens an alert marks it read. It must not become invisible.

    This happened for real during testing: the message was opened to look at its
    attachment, and a read-flag filter made the alarm disappear from the run.
    """
    leido = message()
    outlook = FakeOutlook([leido])  # the fake returns it regardless of read state
    report = build(outlook).run()

    assert report.stored == 1
    assert len(outlook.drafts) == 1


def test_a_pdf_with_a_content_id_is_not_mistaken_for_an_inline_image(build):
    """Gmail sets a Content-ID on every attachment, PDFs included.

    A Content-ID check here once dropped a real report: the alarm was filed as
    "no PDF attached" and no draft was ever created. The only filter is the
    extension, and a signature logo is never a PDF.
    """

    class OutlookConContentId(FakeOutlook):
        def save_pdf_attachments(self, entry_id, dest_dir, max_mb=20):
            # Stands for an attachment that carries a Content-ID: the adapter must
            # still hand it over.
            dest_dir.mkdir(parents=True, exist_ok=True)
            pdf = dest_dir / "clariot original.pdf"
            pdf.write_bytes(b"%PDF-1.4 fake")
            return [pdf]

    outlook = OutlookConContentId([message()])
    report = build(outlook).run()

    assert report.ignored == 0
    assert report.stored == 1
    assert outlook.drafts[0]["attachments"][0].name == "clariot original.pdf"


# --- an open draft grows, and never gets demoted -----------------------------


def test_three_alarms_on_one_machine_grow_a_single_urgent_draft(build):
    """The case the whole project exists for: one pump alarming three times.

    Found by a load test. Two bugs conspired here. The history was pre-filtered by
    "now minus the window" while the classifier windows by the event's own
    timestamp, and the cooldown — which exists so the client is not mailed twice —
    downgraded the third alarm to "normal", renaming the open [URGENTE] draft back
    to a routine report and moving it out of the urgency folder.
    """
    class OutlookConPdfPropio(FakeOutlook):
        """Each alarm brings its own report, as the real ones do."""

        def __init__(self, messages, contenido):
            super().__init__(messages)
            self.contenido = contenido

        def save_pdf_attachments(self, entry_id, dest_dir, max_mb=20):
            dest_dir.mkdir(parents=True, exist_ok=True)
            pdf = dest_dir / "reporte.pdf"
            pdf.write_bytes(self.contenido)
            return [pdf]

    fechas = ["22-07-2026 08:00", "23-07-2026 09:00", "24-07-2026 10:00"]
    outlook = None
    for i, fecha in enumerate(fechas):
        outlook = OutlookConPdfPropio(
            [message(f"<alarma-{i}@aliot>")], f"%PDF-1.4 alarma {i}".encode()
        )
        build(outlook, alert={**ALERT, "event_date": fecha}).run()

    # One draft, updated twice: the client gets one email per machine, not three.
    assert len(outlook.updates) == 1
    ultimo = outlook.updates[-1]
    assert "[URGENTE]" in ultimo["subject"], ultimo["subject"]
    assert ultimo["target_folder"] == "Urgencias"
    # And it carries every report so far, not only the newest.
    assert len(ultimo["attachments"]) == 3, ultimo["attachments"]


def test_the_console_counts_the_draft_as_it_stands_not_as_it_started(build):
    """The summary line is how the technician decides which folder to open.

    The first alarm of a machine is normal and the second promotes it. Counting
    the actions instead of the drafts printed "1 normales" for a run whose only
    draft was sitting in the urgency folder.
    """
    for i, fecha in enumerate(["22-07-2026 08:00", "23-07-2026 09:00"]):
        outlook = FakeOutlook([message(f"<promo-{i}@aliot>")])
        report = build(outlook, alert={**ALERT, "event_date": fecha}).run()

    assert report.urgent_drafts == 1
    assert report.single_drafts == 0
    assert "1 urgentes" in report.summary()
    assert "normales" not in report.summary()


def test_the_reuse_check_is_told_which_copies_are_ours(build):
    """Without the copy list, our own CC reads back as a technician's edit.

    The union is passed on purpose: the severity is only decided after the draft
    has been picked for reuse, so the critical copies count too.
    """
    from clariot.config import ClientDirectory

    build(
        FakeOutlook([message("<cc-1@aliot>")]),
        alert={**ALERT, "event_date": "22-07-2026 08:00"},
    ).run()

    outlook = FakeOutlook([message("<cc-2@aliot>")])
    ingestor = build(outlook, alert={**ALERT, "event_date": "23-07-2026 09:00"})
    ingestor.directory = ClientDirectory(
        routes=[],
        critical_cc=("jefatura@emeltec.cl",),
        always_cc=("archivo@emeltec.cl",),
    )
    ingestor.run()

    assert outlook.state_queries, "the open draft was never checked"
    entry_id, expected_cc = outlook.state_queries[-1]
    assert entry_id == "DRAFT-1"
    assert set(expected_cc) == {"archivo@emeltec.cl", "jefatura@emeltec.cl"}


def test_the_reuse_check_is_told_which_recipients_are_ours(build):
    """A pre-loaded To must not read as the technician's edit.

    Mirrors the copy-list test: clients.yaml fills ``to`` for a known company, so
    the reuse check has to be told how many recipients are ours.
    """
    from clariot.config import ClientDirectory
    from clariot.models import ClientRoute
    from clariot.textutils import normalize

    route = ClientRoute(
        display_name="Soprole",
        to=("uno@soprole.example", "dos@soprole.example"),
        cc=(),
    )
    directory = ClientDirectory(routes=[((normalize("soprole"),), route)])

    first = build(
        FakeOutlook([message("<to-1@aliot>")]),
        alert={**ALERT, "event_date": "22-07-2026 08:00"},
    )
    first.directory = directory
    first.run()

    outlook = FakeOutlook([message("<to-2@aliot>")])
    second = build(outlook, alert={**ALERT, "event_date": "23-07-2026 09:00"})
    second.directory = directory
    second.run()

    assert outlook.to_queries, "the open draft was never checked"
    _, expected_to = outlook.to_queries[-1]
    assert expected_to == ("uno@soprole.example", "dos@soprole.example")


def test_two_different_reports_with_the_same_filename_never_overwrite(build, settings):
    """The "PDF from another email" risk: same name, different content.

    A timestamp suffix alone would collide when several alarms are processed in the
    same second, and one alarm's draft would carry another's PDF.
    """
    contenidos = [b"%PDF-1.4 primero", b"%PDF-1.4 segundo", b"%PDF-1.4 tercero"]

    class OutlookMismoNombre(FakeOutlook):
        def __init__(self, messages):
            super().__init__(messages)
            self.turno = 0

        def save_pdf_attachments(self, entry_id, dest_dir, max_mb=20):
            dest_dir.mkdir(parents=True, exist_ok=True)
            pdf = dest_dir / "reporte.pdf"  # always the same name
            pdf.write_bytes(contenidos[self.turno])
            self.turno += 1
            return [pdf]

    for i, _ in enumerate(contenidos):
        outlook = OutlookMismoNombre([message(f"<n{i}@aliot>")])
        outlook.turno = i
        build(
            outlook,
            alert={**ALERT, "event_date": f"2{i}-07-2026 10:00", "serial_number": f"S-{i}"},
        ).run()
        adjunto = outlook.drafts[0]["attachments"][0]
        # Each draft carries its own bytes, not a neighbour's.
        assert adjunto.read_bytes() == contenidos[i], f"alarma {i} recibio el PDF ajeno"

    archivados = sorted(p.name for p in settings.paths.archive_dir.rglob("*.pdf"))
    assert len(archivados) == 3, archivados
