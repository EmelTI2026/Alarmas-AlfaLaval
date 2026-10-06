"""Domain objects shared across the pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Mapping


@dataclass(frozen=True)
class AlertReport:
    """Key/value data extracted from the original (untranslated) report."""

    fields: Mapping[str, str] = field(default_factory=dict)

    def get(self, name: str) -> str | None:
        value = self.fields.get(name)
        return value or None

    @property
    def company(self) -> str | None:
        return self.get("company")

    @property
    def machine(self) -> str | None:
        return self.get("machine")

    @property
    def urgency(self) -> str | None:
        return self.get("urgency")

    @property
    def machine_label(self) -> str:
        """Machine name, or a neutral placeholder safe to put in a subject line."""
        return self.machine or "Equipo sin identificar"

    @property
    def is_empty(self) -> bool:
        return not any(self.fields.values())


@dataclass(frozen=True)
class ClientRoute:
    """Recipients for one client, resolved from config/clients.yaml."""

    display_name: str
    to: tuple[str, ...]
    cc: tuple[str, ...]
