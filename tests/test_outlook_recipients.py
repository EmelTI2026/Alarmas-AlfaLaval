"""Recipients written on a draft must be resolved, or Outlook cannot send it.

Assigning ``To`` and ``CC`` as text leaves every entry unresolved: no address,
shown in red, and sending fails until a person picks each one by hand. A real
Outlook confirmed it: all six recipients read back ``Resolved = False`` until
``Recipients.ResolveAll()`` ran.

No real Outlook is involved, so the instance is built without COM and handed a
fake mail item that records what was done to it, in order.
"""

from pathlib import Path
from types import SimpleNamespace

import pytest

from clariot.adapters.outlook import OutlookClient


class FakeRecipients:
    def __init__(self, ops: list[str], *, resolves: bool = True, error: str = "") -> None:
        self._ops = ops
        self._resolves = resolves
        self._error = error
        self.names: list[str] = []
        self.resolved: set[str] = set()

    @property
    def Count(self) -> int:
        return len(self.names)

    def Item(self, index: int):
        name = self.names[index - 1]
        return SimpleNamespace(Name=name, Resolved=name in self.resolved)

    def ResolveAll(self) -> bool:
        self._ops.append("resolve")
        if self._error:
            raise RuntimeError(self._error)
        if self._resolves:
            self.resolved = set(self.names)
        else:
            # One address stays unknown, as a mistyped one would.
            self.resolved = set(self.names[:-1])
        return self._resolves


class FakeAttachments:
    Count = 0

    def Add(self, path: str) -> None:  # pragma: no cover - never reached here
        raise AssertionError("no attachment expected")

    def Item(self, index: int):  # pragma: no cover
        raise AssertionError("no attachment expected")


class FakeMail:
    EntryID = "DRAFT-1"

    def __init__(self, *, resolves: bool = True, error: str = "") -> None:
        self.ops: list[str] = []
        self.Recipients = FakeRecipients(self.ops, resolves=resolves, error=error)
        self.Attachments = FakeAttachments()
        self.Subject = ""
        self.HTMLBody = ""
        self.Sent = False
        self.saves = 0
        self.Parent = SimpleNamespace(Name="Por enviar")
        self._to = ""
        self._cc = ""

    @property
    def To(self) -> str:
        return self._to

    @To.setter
    def To(self, value: str) -> None:
        self._to = value
        self.ops.append("to")
        self.Recipients.names = [*self.Recipients.names, *value.split("; ")]

    @property
    def CC(self) -> str:
        return self._cc

    @CC.setter
    def CC(self, value: str) -> None:
        self._cc = value
        self.ops.append("cc")
        # Assigning text replaces the CC entries with unresolved ones.
        self.Recipients.names = [*self.Recipients.names, *value.split("; ")]
        self.Recipients.resolved = set()

    def Save(self) -> None:
        self.saves += 1
        self.ops.append("save")


def client(mail: FakeMail) -> OutlookClient:
    adapter = object.__new__(OutlookClient)
    adapter._app = SimpleNamespace(CreateItem=lambda _kind: mail)
    adapter._namespace = SimpleNamespace(GetItemFromID=lambda entry_id: mail)
    adapter._use_signature = False
    adapter._signature_cache = None
    adapter._signatures_dir = Path("no-signatures-here")
    adapter._signature_account = ""
    return adapter


TO = ("rletelier@koandina.com", "mvalenzuelal@koandina.com")
CC = ("porellana@emeltec.cl", "cfarias@emeltec.cl")


def create(adapter: OutlookClient, **kwargs) -> str:
    return adapter.create_draft(
        subject="[URGENTE] Bomba",
        html_body="<p>informe</p>",
        to=kwargs.get("to", TO),
        cc=kwargs.get("cc", CC),
        attachments=(),
    )


def test_every_recipient_of_a_new_draft_is_resolved():
    """The regression: a draft left in red that Outlook would not send."""
    mail = FakeMail()

    create(client(mail))

    assert mail.Recipients.resolved == set(mail.Recipients.names)
    assert len(mail.Recipients.names) == 4


def test_recipients_are_resolved_after_they_are_written_and_before_saving():
    mail = FakeMail()

    create(client(mail))

    ops = mail.ops
    assert ops.index("resolve") > max(ops.index("to"), ops.index("cc"))
    assert ops.index("resolve") < ops.index("save")


def test_a_draft_without_recipients_is_still_created():
    """Unknown clients get an empty To, and the draft must not depend on it."""
    mail = FakeMail()

    create(client(mail), to=(), cc=())

    assert mail.saves == 1


def test_an_unresolvable_address_is_reported_and_the_draft_is_still_saved(caplog):
    mail = FakeMail(resolves=False)

    with caplog.at_level("WARNING"):
        entry_id = create(client(mail))

    assert entry_id == "DRAFT-1"
    assert mail.saves == 1
    assert "could not resolve these recipients" in caplog.text
    assert "cfarias@emeltec.cl" in caplog.text


def test_a_failure_while_resolving_never_blocks_the_draft(caplog):
    mail = FakeMail(error="Outlook is busy")

    with caplog.at_level("WARNING"):
        entry_id = create(client(mail))

    assert entry_id == "DRAFT-1"
    assert mail.saves == 1
    assert "Could not resolve the recipients" in caplog.text


def test_growing_a_draft_resolves_the_copies_it_rewrites():
    """Writing CC again as text drops the resolution the first write earned."""
    mail = FakeMail()
    adapter = client(mail)
    create(adapter)
    mail.ops.clear()

    adapter.update_draft("DRAFT-1", subject="dos alarmas", html_body="<p>x</p>", cc=CC)

    assert mail.ops.index("cc") < mail.ops.index("resolve") < mail.ops.index("save")
    assert set(mail.Recipients.names[-len(CC):]) <= mail.Recipients.resolved


def test_growing_a_draft_without_copies_leaves_recipients_alone():
    mail = FakeMail()
    adapter = client(mail)
    create(adapter)
    mail.ops.clear()

    adapter.update_draft("DRAFT-1", subject="dos alarmas", html_body="<p>x</p>")

    assert "resolve" not in mail.ops
