"""Configuration loading: YAML for behaviour, environment for secrets."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

import yaml
from dotenv import load_dotenv

from .models import ClientRoute
from .textutils import normalize


class ConfigError(RuntimeError):
    """Raised when configuration is missing or internally inconsistent."""


@dataclass(frozen=True)
class OutlookSettings:
    """Folders and mailbox behaviour."""
    source_folders: tuple[str, ...]
    include_subfolders: bool
    processed_folder: str
    draft_folder: str
    urgent_draft_folder: str
    draft_folder_parent: str
    mark_as_read: bool
    max_attachment_mb: int


@dataclass(frozen=True)
class TranslationSettings:
    """Who translates the attached PDF, and into what."""
    provider: str
    target_lang: str
    source_lang: str
    google_location: str
    formality: str
    attach_original: bool

    @property
    def enabled(self) -> bool:
        return self.provider != "none"


@dataclass(frozen=True)
class GlossarySettings:
    """Who translates the report phrases, and what to do with a gap."""
    on_missing: str
    never_hold_urgent: bool
    phrase_provider: str
    gemini_model: str

    @property
    def holds(self) -> bool:
        return self.on_missing == "hold"


@dataclass(frozen=True)
class GroupingSettings:
    """The thresholds that turn repeated alarms into an urgency."""
    window_days: int
    same_day_is_critical: bool
    urgent_threshold: int
    urgency_cooldown_days: int
    retention_days: int


@dataclass(frozen=True)
class ReportSettings:
    """Subject lines, and whether isolated alarms wait for Friday."""
    single_alarm: str
    critical_subject: str
    urgent_subject: str
    normal_subject: str
    weekly_subject: str
    include_emeltec_note: bool
    brand_color: str = "#005f7f"
    """Hex colour of the section headers. Set it to the company's own.

    Defaulted rather than required: it is presentation, so a caller that does not
    care about branding must not have to supply it.
    """

    @property
    def single_alarm_as_draft(self) -> bool:
        return self.single_alarm == "draft"


@dataclass(frozen=True)
class PathSettings:
    """Where the state, the archive and the log live."""
    work_dir: Path
    archive_dir: Path
    ledger_db: Path
    audit_csv: Path
    events_db: Path
    log_file: Path


@dataclass(frozen=True)
class EmailSettings:
    """Greeting, placeholders and the sending account."""
    subject_template: str
    critical_subject_prefix: str
    unknown_client_subject_prefix: str
    greeting_placeholder: str
    greeting: str
    include_report: bool
    sender_account: str
    signature_team: str
    use_outlook_signature: bool


@dataclass(frozen=True)
class Settings:
    """Everything from settings.yaml plus the secrets from .env."""
    project_root: Path
    outlook: OutlookSettings
    translation: TranslationSettings
    glossary: GlossarySettings
    grouping: GroupingSettings
    reports: ReportSettings
    paths: PathSettings
    email: EmailSettings
    critical_urgencies: frozenset[str]
    deepl_api_key: str
    allow_free_deepl_key: bool
    google_project_id: str
    gemini_api_key: str


def _section(data: Mapping[str, Any], name: str) -> Mapping[str, Any]:
    value = data.get(name)
    if not isinstance(value, Mapping):
        raise ConfigError(f"settings.yaml is missing the '{name}' section")
    return value


def _folder_list(outlook: Mapping[str, Any]) -> tuple[str, ...]:
    """Read outlook.source_folders, accepting a bare string for convenience."""
    raw = outlook.get("source_folders", outlook.get("source_folder"))
    if isinstance(raw, str):
        raw = [raw]
    folders = tuple(str(item).strip() for item in (raw or []) if str(item).strip())
    if not folders:
        raise ConfigError(
            "settings.yaml needs at least one entry under outlook.source_folders"
        )
    return folders


def _resolve(root: Path, value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else root / path


def load_settings(project_root: Path, config_dir: Path | None = None) -> Settings:
    """Read config/settings.yaml plus the .env file at the project root."""
    config_dir = config_dir or project_root / "config"
    settings_path = config_dir / "settings.yaml"
    if not settings_path.exists():
        raise ConfigError(f"Missing configuration file: {settings_path}")

    data = yaml.safe_load(settings_path.read_text(encoding="utf-8")) or {}
    load_dotenv(project_root / ".env")

    outlook = _section(data, "outlook")
    translation = _section(data, "translation")
    paths = _section(data, "paths")
    email = _section(data, "email")
    urgency = _section(data, "urgency")
    # Optional sections: absent means defaults, so old config files keep working.
    grouping = data.get("grouping") or {}
    glossary_cfg = data.get("glossary") or {}
    reports = data.get("reports") or {}

    source_folders = _folder_list(outlook)
    processed = str(outlook.get("processed_folder", "")).strip()
    if processed and any(
        processed.lower().startswith(source.lower()) for source in source_folders
    ):
        # Otherwise processed mail would sit inside the tree that gets scanned.
        raise ConfigError(
            f"outlook.processed_folder ('{processed}') must not live inside a "
            "source folder, or processed alerts would be scanned again"
        )

    api_key = os.getenv("DEEPL_API_KEY", "").strip()
    allow_free = os.getenv("DEEPL_ALLOW_FREE_KEY", "false").strip().lower() in {
        "1",
        "true",
        "yes",
    }

    return Settings(
        project_root=project_root,
        outlook=OutlookSettings(
            source_folders=source_folders,
            include_subfolders=bool(outlook.get("include_subfolders", True)),
            processed_folder=processed,
            draft_folder=str(outlook.get("draft_folder", "")).strip(),
            urgent_draft_folder=str(outlook.get("urgent_draft_folder", "")).strip(),
            draft_folder_parent=str(outlook.get("draft_folder_parent", "drafts"))
            .strip()
            .lower(),
            mark_as_read=bool(outlook.get("mark_as_read", True)),
            max_attachment_mb=int(outlook.get("max_attachment_mb", 20)),
        ),
        translation=TranslationSettings(
            provider=str(translation.get("provider", "none")).strip().lower(),
            source_lang=str(translation.get("source_lang", "")).strip().lower(),
            google_location=str(
                (translation.get("google") or {}).get("location", "us-central1")
            ).strip(),
            target_lang=str(translation.get("target_lang", "ES")).strip().upper(),
            formality=str(translation.get("formality", "default")).strip().lower(),
            attach_original=bool(translation.get("attach_original", True)),
        ),
        glossary=GlossarySettings(
            on_missing=str(glossary_cfg.get("on_missing", "hold")).strip().lower(),
            never_hold_urgent=bool(glossary_cfg.get("never_hold_urgent", True)),
            phrase_provider=str(glossary_cfg.get("phrase_provider", "none")).strip().lower(),
            gemini_model=str(glossary_cfg.get("gemini_model", "gemini-flash-latest")).strip(),
        ),
        grouping=GroupingSettings(
            window_days=int(grouping.get("window_days", 7)),
            same_day_is_critical=bool(grouping.get("same_day_is_critical", True)),
            urgent_threshold=int(grouping.get("urgent_threshold", 2)),
            urgency_cooldown_days=int(grouping.get("urgency_cooldown_days", 7)),
            retention_days=int(grouping.get("retention_days", 180)),
        ),
        reports=ReportSettings(
            single_alarm=str(reports.get("single_alarm", "draft")).strip().lower(),
            normal_subject=str(
                reports.get("normal_subject", "Reporte de monitoreo - {machine} - {date}")
            ),
            critical_subject=str(reports.get("critical_subject", "[CRITICO] {machine}")),
            urgent_subject=str(reports.get("urgent_subject", "[URGENTE] {machine}")),
            weekly_subject=str(reports.get("weekly_subject", "Informe semanal - {company}")),
            include_emeltec_note=bool(reports.get("include_emeltec_note", True)),
            brand_color=str(reports.get("brand_color", "#005f7f")).strip() or "#005f7f",
        ),
        paths=PathSettings(
            work_dir=_resolve(project_root, str(paths.get("work_dir", ".work"))),
            archive_dir=_resolve(project_root, str(paths.get("archive_dir", "archive"))),
            ledger_db=_resolve(project_root, str(paths.get("ledger_db", "state/ledger.db"))),
            audit_csv=_resolve(project_root, str(paths.get("audit_csv", "archive/audit.csv"))),
            events_db=_resolve(project_root, str(paths.get("events_db", "state/events.db"))),
            log_file=_resolve(project_root, str(paths.get("log_file", "logs/clariot.log"))),
        ),
        email=EmailSettings(
            subject_template=str(email.get("subject_template", "{machine}")),
            critical_subject_prefix=str(email.get("critical_subject_prefix", "")),
            unknown_client_subject_prefix=str(email.get("unknown_client_subject_prefix", "")),
            greeting_placeholder=str(email.get("greeting_placeholder", "[NOMBRE]")),
            greeting=str(email.get("greeting", "Buenos días")),
            include_report=bool(email.get("include_report", True)),
            sender_account=str(email.get("sender_account", "")).strip(),
            signature_team=str(email.get("signature_team", "")).strip(),
            use_outlook_signature=bool(email.get("use_outlook_signature", True)),
        ),
        critical_urgencies=frozenset(
            normalize(str(item)) for item in urgency.get("critical_values", [])
        ),
        deepl_api_key=api_key,
        allow_free_deepl_key=allow_free,
        gemini_api_key=os.getenv("GEMINI_API_KEY", "").strip(),
        google_project_id=os.getenv("GOOGLE_CLOUD_PROJECT", "").strip()
        or str((translation.get("google") or {}).get("project_id", "")).strip(),
    )


def load_labels(config_dir: Path) -> dict[str, tuple[str, ...]]:
    """Read the canonical-field -> label-aliases map, longest alias first."""
    labels_path = config_dir / "pdf_labels.yaml"
    if not labels_path.exists():
        raise ConfigError(f"Missing configuration file: {labels_path}")

    data = yaml.safe_load(labels_path.read_text(encoding="utf-8")) or {}
    fields = data.get("fields")
    if not isinstance(fields, Mapping):
        raise ConfigError("pdf_labels.yaml is missing the 'fields' section")

    result: dict[str, tuple[str, ...]] = {}
    for canonical, aliases in fields.items():
        if not isinstance(aliases, Sequence) or isinstance(aliases, str):
            raise ConfigError(f"Field '{canonical}' must map to a list of labels")
        # Longest first, so "MACHINE NAME" wins over the "MACHINE" prefix.
        result[str(canonical)] = tuple(
            sorted((str(a) for a in aliases), key=len, reverse=True)
        )
    return result


def load_value_noise(config_dir: Path) -> tuple[str, ...]:
    """Fragments stripped from extracted values (help links glued to a value)."""
    labels_path = config_dir / "pdf_labels.yaml"
    data = yaml.safe_load(labels_path.read_text(encoding="utf-8")) or {}
    return tuple(str(item) for item in data.get("value_noise") or ())


class ClientDirectory:
    """Resolves an extracted company name to its recipients."""

    def __init__(
        self,
        routes: Sequence[tuple[tuple[str, ...], ClientRoute]],
        critical_cc: Sequence[str] = (),
        always_cc: Sequence[str] = (),
    ) -> None:
        # Longest patterns first so a plant-specific rule beats a corporate one.
        self._routes = sorted(
            routes, key=lambda item: max((len(p) for p in item[0]), default=0), reverse=True
        )
        self.critical_cc = tuple(critical_cc)
        self.always_cc = tuple(always_cc)

    @classmethod
    def load(cls, config_dir: Path) -> "ClientDirectory":
        clients_path = config_dir / "clients.yaml"
        if not clients_path.exists():
            raise ConfigError(f"Missing configuration file: {clients_path}")

        data = yaml.safe_load(clients_path.read_text(encoding="utf-8")) or {}
        routes: list[tuple[tuple[str, ...], ClientRoute]] = []
        for entry in data.get("clients") or []:
            patterns = entry.get("match") or []
            if isinstance(patterns, str):
                patterns = [patterns]
            recipients = tuple(entry.get("to") or ())
            if not patterns or not recipients:
                raise ConfigError(
                    f"Client entry {entry!r} needs at least one 'match' and one 'to' address"
                )
            routes.append(
                (
                    tuple(normalize(str(p)) for p in patterns),
                    ClientRoute(
                        display_name=str(entry.get("display_name") or patterns[0]).strip(),
                        to=recipients,
                        cc=tuple(entry.get("cc") or ()),
                    ),
                )
            )
        return cls(
            routes,
            critical_cc=tuple(data.get("critical_cc") or ()),
            always_cc=tuple(data.get("always_cc") or ()),
        )

    def resolve(self, company: str | None) -> ClientRoute | None:
        """Return the route whose pattern appears in the company field, if any."""
        if not company:
            return None
        haystack = normalize(company)
        for patterns, route in self._routes:
            if any(pattern and pattern in haystack for pattern in patterns):
                return route
        return None

    def __len__(self) -> int:
        return len(self._routes)
