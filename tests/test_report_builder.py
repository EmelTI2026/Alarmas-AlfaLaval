from dataclasses import replace
from datetime import datetime
from pathlib import Path

import pytest

from clariot.aggregate import group_by_equipment, summarize_equipment
from clariot.classifier import LEVEL_CRITICAL, LEVEL_URGENT
from clariot.config import ClientDirectory
from clariot.report_builder import build_template_env
from clariot.glossary import Glossary
from clariot.models import ClientRoute
from clariot.report_builder import build_equipment_draft, build_weekly_draft
from clariot.store import (
    REPORT_CRITICAL,
    REPORT_URGENT,
    REPORT_WEEKLY,
    StoredEvent,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def event(when, *, serial="VX-3037575", company="Soprole", event_type="Mechanical", pdf=""):
    return StoredEvent(
        event_key=f"{company}|{serial}|{when.isoformat()}",
        equipment_key=f"{company.upper()}|{serial}",
        company=company,
        plant=f"{company} - San Bernardo",
        machine="VX-RetCIPLin1 1206230",
        serial_number=serial,
        machine_type="Centrifugal_pump",
        event_at=when,
        event_date_raw=when.strftime("%d-%m-%Y %H:%M"),
        event_type=event_type,
        equipment_status="Potential impeller damage or imbalance.",
        possible_cause="A foreign object may have passed through the pump.",
        recommended_action="Open pump and check impeller for any damage.",
        urgency="Perform check at next planned stop of the pump.",
        pdf_path=pdf,
        pdf_translated="",
        message_key="<x@clariot>",
        ingested_at=when.isoformat(),
        reported_at=None,
        report_kind="",
    )


@pytest.fixture
def env():
    return build_template_env(PROJECT_ROOT / "templates")


@pytest.fixture
def glossary():
    return Glossary(
        {
            "Mechanical": "Mecánico",
            "Centrifugal_pump": "Bomba centrífuga",
            "Potential impeller damage or imbalance.": "Posible daño del impulsor.",
            "Open pump and check impeller for any damage.": "Abrir la bomba y revisar el impulsor.",
            "Perform check at next planned stop of the pump.": "Revisar en la próxima parada.",
        }
    )


@pytest.fixture
def soprole():
    return ClientDirectory(
        routes=[
            (("SOPROLE",), ClientRoute("Soprole", ("mantencion@soprole.example",), ()))
        ],
        critical_cc=("jefatura@emeltec.cl",),
    )


# --- urgency ----------------------------------------------------------------


def test_a_critical_report_names_the_count_and_the_day(settings, soprole, env, glossary):
    events = [event(datetime(2026, 7, 26, 9, 15)), event(datetime(2026, 7, 26, 14, 2))]
    summary = summarize_equipment(events, glossary, settings.critical_urgencies)

    draft = build_equipment_draft(summary, LEVEL_CRITICAL, soprole, settings, env)

    assert draft.subject == "[CRITICO] VX-RetCIPLin1 1206230 - 2 alarmas el 26 de julio"
    assert draft.kind == REPORT_CRITICAL
    assert draft.critical is True


def test_a_critical_report_copies_the_escalation_list(settings, soprole, env, glossary):
    summary = summarize_equipment([event(datetime(2026, 7, 26, 9, 0))], glossary)
    draft = build_equipment_draft(summary, LEVEL_CRITICAL, soprole, settings, env)

    assert draft.to == ("mantencion@soprole.example",)
    assert "jefatura@emeltec.cl" in draft.cc


def test_an_urgent_report_does_not_copy_the_escalation_list(settings, soprole, env, glossary):
    summary = summarize_equipment([event(datetime(2026, 7, 26, 9, 0))], glossary)
    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert draft.kind == REPORT_URGENT
    assert draft.cc == ()


def test_the_whole_chronology_appears_translated(settings, soprole, env, glossary):
    events = [
        event(datetime(2026, 7, 24, 22, 34)),
        event(datetime(2026, 7, 26, 14, 2)),
    ]
    summary = summarize_equipment(events, glossary, settings.critical_urgencies)
    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert "24 de julio, 22:34" in draft.html_body
    assert "26 de julio, 14:02" in draft.html_body
    assert "Mecánico" in draft.html_body
    # No English leaks into the chronology.
    assert "Mechanical" not in draft.html_body
    assert "Perform check at next planned stop" not in draft.html_body


def test_every_block_of_the_email_has_the_same_width(settings, soprole, env, glossary):
    """Seen in a real email: each block collapsed to its own content width.

    The EQUIPO card came out narrow, the alarm card wide and the recommendation
    narrower still, which reads as broken rather than as a design. The Word engine
    honours the HTML ``width`` attribute, so every table has to declare it.
    """
    import re

    events = [
        replace(event(datetime(2026, 7, 22)), event_type="Mechanical"),
        replace(
            event(datetime(2026, 7, 24)),
            event_type="Bearing",
            equipment_status="Bearing wear.",
        ),
    ]
    summary = summarize_equipment(events, glossary)

    for nivel in (LEVEL_URGENT, LEVEL_CRITICAL):
        draft = build_equipment_draft(summary, nivel, soprole, settings, env)
        anchos = re.findall(r"<table[^>]*\swidth=\"(\d+)\"", draft.html_body)
        assert len(anchos) >= 4, f"{nivel}: faltan tablas con ancho declarado"
        assert len(set(anchos)) == 1, f"{nivel}: anchos distintos {set(anchos)}"


def test_the_brand_colour_comes_from_the_configuration(settings, soprole, env, glossary):
    """One line rebrands both reports. Hard-coded colours cannot be changed."""
    marca = replace(settings, reports=replace(settings.reports, brand_color="#0d5c3f"))
    summary = summarize_equipment([event(datetime(2026, 7, 26))], glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, marca, env)

    assert "#0d5c3f" in draft.html_body
    assert "#005f7f" not in draft.html_body  # el corporativo por omisión ya no aparece


def test_every_distinct_condition_gets_its_own_detail(settings, soprole, env, glossary):
    """A pump reporting looseness and then a bearing must show both.

    The body used to carry only the latest alarm's condition and action, so the
    earlier failure disappeared from the email entirely whenever the two differed.
    """
    holgura = replace(
        event(datetime(2026, 7, 22, 6, 12)),
        event_type="Mechanical",
        equipment_status="Looseness or installation issue.",
        recommended_action="Check the fixing bolts and the base.",
    )
    rodamiento = replace(
        event(datetime(2026, 7, 24, 21, 5)),
        event_type="Bearing",
        equipment_status="Bearing wear on the drive end.",
        recommended_action="Replace the drive end bearing.",
    )
    summary = summarize_equipment([holgura, rodamiento], glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert "2 condiciones distintas" in draft.html_body
    for texto in (
        "Looseness or installation issue.",
        "Check the fixing bolts and the base.",
        "Bearing wear on the drive end.",
        "Replace the drive end bearing.",
    ):
        assert texto in draft.html_body, texto


def test_one_condition_repeated_is_written_once_with_all_its_dates(
    settings, soprole, env, glossary
):
    """Three identical reports must not print the same paragraph three times."""
    events = [
        event(datetime(2026, 7, 22, 6, 12)),
        event(datetime(2026, 7, 23, 14, 48)),
        event(datetime(2026, 7, 24, 21, 5)),
    ]
    summary = summarize_equipment(events, glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert draft.html_body.count("Posible daño del impulsor.") == 1
    assert "3 alarmas · 22, 23 y 24 de julio" in draft.html_body
    assert "condiciones distintas" not in draft.html_body


def test_a_condition_already_sent_does_not_come_back_in_the_detail(
    settings, soprole, env, glossary
):
    """Once the draft went out, its condition must not be described again."""
    enviada = replace(
        event(datetime(2026, 7, 22, 6, 12)),
        equipment_status="Looseness or installation issue.",
        recommended_action="Check the fixing bolts and the base.",
        reported_at="2026-07-22T10:00:00",
        report_kind=REPORT_URGENT,
    )
    nueva = replace(
        event(datetime(2026, 7, 24, 21, 5)),
        event_type="Bearing",
        equipment_status="Bearing wear on the drive end.",
        recommended_action="Replace the drive end bearing.",
    )
    summary = summarize_equipment([enviada, nueva], glossary)

    draft = build_equipment_draft(summary, LEVEL_CRITICAL, soprole, settings, env)

    assert "Bearing wear on the drive end." in draft.html_body
    assert "Looseness or installation issue." not in draft.html_body
    assert "Check the fixing bolts and the base." not in draft.html_body
    # But the client still learns why the severity is what it is.
    assert "La primera ya fue informada" in draft.html_body


def test_every_event_is_reported_so_none_is_counted_again(settings, soprole, env, glossary):
    events = [event(datetime(2026, 7, 24)), event(datetime(2026, 7, 26))]
    summary = summarize_equipment(events, glossary)
    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert set(draft.event_keys) == {e.event_key for e in events}


def test_alarms_already_sent_are_not_sent_again(settings, soprole, env, glossary):
    """The client received the earlier ones; repeating them looks like a mistake."""
    sent = replace(
        event(datetime(2026, 7, 24, 22, 34)),
        reported_at="2026-07-26T10:00:00",
        report_kind=REPORT_URGENT,
    )
    fresh = event(datetime(2026, 7, 26, 14, 2))
    summary = summarize_equipment([sent, fresh], glossary)

    draft = build_equipment_draft(summary, LEVEL_CRITICAL, soprole, settings, env)

    # Only the new alarm is detailed, attached and marked.
    assert draft.event_keys == (fresh.event_key,)
    assert "26 de julio, 14:02" in draft.html_body
    assert "24 de julio, 22:34" not in draft.html_body


def test_the_accumulated_count_still_explains_the_severity(settings, soprole, env, glossary):
    """Showing one alarm under a CRITICO subject would confuse the client."""
    sent = [
        replace(
            event(datetime(2026, 7, 24, 22, 34)),
            reported_at="2026-07-26T10:00:00",
            report_kind=REPORT_URGENT,
        ),
        replace(
            event(datetime(2026, 7, 26, 9, 15)),
            reported_at="2026-07-26T10:00:00",
            report_kind=REPORT_URGENT,
        ),
    ]
    summary = summarize_equipment([*sent, event(datetime(2026, 7, 26, 14, 2))], glossary)

    draft = build_equipment_draft(summary, LEVEL_CRITICAL, soprole, settings, env)

    assert "acumula <strong>3 alarmas</strong>" in draft.html_body
    assert "Las 2 anteriores ya fueron informadas" in draft.html_body
    assert "1 ALARMA NUEVA" in draft.html_body
    # The subject keeps the total, because that is the severity signal.
    assert "3 alarmas" in draft.subject


def test_a_first_report_shows_everything_and_no_context_note(settings, soprole, env, glossary):
    events = [event(datetime(2026, 7, 24)), event(datetime(2026, 7, 26))]
    summary = summarize_equipment(events, glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert "ya fueron informadas" not in draft.html_body
    assert "ya fue informada" not in draft.html_body
    assert len(draft.event_keys) == 2


def test_the_emeltec_block_can_be_switched_off(settings, soprole, env, glossary):
    off = replace(settings, reports=replace(settings.reports, include_emeltec_note=False))
    summary = summarize_equipment([event(datetime(2026, 7, 26))], glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, off, env)

    assert "RECOMENDACIÓN DE EMELTEC" not in draft.html_body
    assert "[COMPLETAR]" not in draft.html_body


def test_an_unknown_client_still_gets_a_draft(settings, env, glossary):
    empty = ClientDirectory(routes=[])
    summary = summarize_equipment([event(datetime(2026, 7, 26))], glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, empty, settings, env)

    assert draft.to == ()
    assert draft.client_resolved is False
    assert "[NOMBRE]" in draft.html_body


def test_the_greeting_keeps_the_name_placeholder_for_a_known_client(
    settings, soprole, env, glossary
):
    """The greeting goes to the person, so a known company must not replace it."""
    summary = summarize_equipment([event(datetime(2026, 7, 26))], glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert draft.client_resolved is True
    assert draft.to == ("mantencion@soprole.example",)
    assert f"{settings.email.greeting} [NOMBRE]," in draft.html_body
    assert f"{settings.email.greeting} Soprole," not in draft.html_body


def test_the_weekly_greeting_keeps_the_name_placeholder_for_a_known_client(
    settings, soprole, env, glossary
):
    summaries = group_by_equipment([event(datetime(2026, 7, 22))], glossary)

    draft = build_weekly_draft("Soprole", summaries, soprole, settings, env)

    assert f"{settings.email.greeting} [NOMBRE]," in draft.html_body
    assert f"{settings.email.greeting} Soprole," not in draft.html_body


def test_a_broken_subject_template_says_which_field_is_wrong(settings, soprole, env, glossary):
    broken = replace(settings, reports=replace(settings.reports, urgent_subject="{inventado}"))
    summary = summarize_equipment([event(datetime(2026, 7, 26))], glossary)

    with pytest.raises(ValueError, match="campo desconocido"):
        build_equipment_draft(summary, LEVEL_URGENT, soprole, broken, env)


def test_missing_attachment_files_are_skipped(settings, soprole, env, glossary, tmp_path):
    real = tmp_path / "reporte.pdf"
    real.write_bytes(b"%PDF-1.4")
    events = [
        event(datetime(2026, 7, 26), pdf=str(real)),
        event(datetime(2026, 7, 25), pdf=str(tmp_path / "borrado.pdf")),
    ]
    summary = summarize_equipment(events, glossary)

    draft = build_equipment_draft(summary, LEVEL_URGENT, soprole, settings, env)

    assert draft.attachments == (real,)


# --- weekly -----------------------------------------------------------------


def test_the_weekly_report_separates_repeated_from_isolated(settings, soprole, env, glossary):
    events = [
        event(datetime(2026, 7, 22), serial="B-2"),
        event(datetime(2026, 7, 23), serial="C-3"),
        event(datetime(2026, 7, 24), serial="C-3"),
    ]
    summaries = group_by_equipment(events, glossary, settings.critical_urgencies)

    draft = build_weekly_draft("Soprole", summaries, soprole, settings, env)

    assert draft.kind == REPORT_WEEKLY
    assert "alarmas reiteradas" in draft.html_body
    assert "una alarma en el período" in draft.html_body
    assert "3 alarmas en" in draft.html_body


def test_the_weekly_subject_carries_client_and_period(settings, soprole, env, glossary):
    events = [event(datetime(2026, 7, 21)), event(datetime(2026, 7, 27), serial="B-2")]
    summaries = group_by_equipment(events, glossary)

    draft = build_weekly_draft("Soprole", summaries, soprole, settings, env)

    assert draft.subject == (
        "Informe semanal de monitoreo - Soprole - 21 al 27 de julio"
    )


def test_a_weekly_report_with_only_isolated_alarms_omits_the_repeat_table(
    settings, soprole, env, glossary
):
    summaries = group_by_equipment([event(datetime(2026, 7, 22))], glossary)
    draft = build_weekly_draft("Soprole", summaries, soprole, settings, env)

    assert "alarmas reiteradas" not in draft.html_body
    assert "una alarma en el período" in draft.html_body


def test_the_weekly_report_reports_every_event(settings, soprole, env, glossary):
    events = [event(datetime(2026, 7, 22 + i), serial=f"S-{i}") for i in range(3)]
    summaries = group_by_equipment(events, glossary)

    draft = build_weekly_draft("Soprole", summaries, soprole, settings, env)

    assert len(draft.event_keys) == 3


def test_an_empty_weekly_report_is_rejected(settings, soprole, env):
    with pytest.raises(ValueError):
        build_weekly_draft("Soprole", [], soprole, settings, env)
