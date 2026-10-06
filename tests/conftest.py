from pathlib import Path

import pytest

from clariot.config import (
    ClientDirectory,
    GlossarySettings,
    GroupingSettings,
    ReportSettings,
    EmailSettings,
    OutlookSettings,
    PathSettings,
    Settings,
    TranslationSettings,
)
from clariot.report_builder import build_template_env
from clariot.models import ClientRoute

PROJECT_ROOT = Path(__file__).resolve().parents[1]

LABELS = {
    "company": ("COMPANY", "COMPANIA", "PLANTA", "PLANT"),
    "machine": ("NOMBRE DE LA MAQUINA", "MACHINE NAME", "MACHINE"),
    "sensor_id": ("SENSOR ID", "ID DEL SENSOR"),
    "event_type": ("EVENT TYPE", "TIPO DE EVENTO"),
    "urgency": ("URGENCY", "URGENCIA"),
    "event_date": ("EVENT DATE", "FECHA"),
    "location": ("LOCATION", "UBICACION"),
}


@pytest.fixture
def labels():
    return LABELS


@pytest.fixture
def settings(tmp_path):
    return Settings(
        project_root=tmp_path,
        outlook=OutlookSettings(
            source_folders=("1_Auditoria_Clariot",),
            include_subfolders=True,
            processed_folder="2_Auditoria_Procesado",
            draft_folder="3_Por_Enviar",
            urgent_draft_folder="Urgencias",
            draft_folder_parent="drafts",
            mark_as_read=True,
            max_attachment_mb=20,
        ),
        translation=TranslationSettings(
            provider="google",
            target_lang="ES",
            source_lang="",
            google_location="us-central1",
            formality="more",
            attach_original=True,
        ),
        glossary=GlossarySettings(
            on_missing="hold",
            never_hold_urgent=True,
            phrase_provider="none",
            gemini_model="gemini-flash-latest",
        ),
        grouping=GroupingSettings(
            window_days=7,
            same_day_is_critical=True,
            urgent_threshold=2,
            urgency_cooldown_days=7,
            retention_days=180,
        ),
        reports=ReportSettings(
            single_alarm="draft",
            normal_subject="Reporte de monitoreo - {machine} - {date}",
            critical_subject="[CRITICO] {machine} - {count} alarmas el {date}",
            urgent_subject="[URGENTE] {machine} - {count} alarmas ({period})",
            weekly_subject="Informe semanal de monitoreo - {company} - {period}",
            include_emeltec_note=True,
        ),
        paths=PathSettings(
            work_dir=tmp_path / "work",
            archive_dir=tmp_path / "archive",
            ledger_db=tmp_path / "state" / "ledger.db",
            audit_csv=tmp_path / "archive" / "audit.csv",
            events_db=tmp_path / "state" / "events.db",
            log_file=tmp_path / "logs" / "clariot.log",
        ),
        email=EmailSettings(
            subject_template="[ALERTA TECNICA] Reporte de Estado - {machine}",
            critical_subject_prefix="[CRITICO] ",
            unknown_client_subject_prefix="",
            greeting_placeholder="[NOMBRE]",
            greeting="Buenos días",
            include_report=True,
            sender_account="",
            signature_team="Equipo Emeltec Chile",
            use_outlook_signature=True,
        ),
        critical_urgencies=frozenset({"CRITICAL", "CRITICA", "ALTA"}),
        deepl_api_key="test-key",
        allow_free_deepl_key=False,
        google_project_id="proyecto-de-prueba",
        gemini_api_key="clave-de-prueba",
    )


@pytest.fixture
def directory():
    return ClientDirectory(
        routes=[
            (("NESTLE",), ClientRoute("Nestlé", ("planta@nestle.example",), ())),
            (("PROLESUR",), ClientRoute("Prolesur", ("ops@prolesur.example",), ("sup@emeltec.cl",))),
        ],
        critical_cc=("jefatura@emeltec.cl",),
        always_cc=(),
    )


@pytest.fixture
def env():
    return build_template_env(PROJECT_ROOT / "templates")
