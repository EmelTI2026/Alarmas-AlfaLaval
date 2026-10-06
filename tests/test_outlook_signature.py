"""The draft is signed by Outlook itself, so nobody pastes a signature by hand.

No real Outlook is involved. ``OutlookClient.__init__`` dispatches COM, which
only exists on a Windows machine running classic Outlook, so the instance is
built without it and handed a fake application that behaves the way Outlook
does: touching ``GetInspector`` on a mail item drops the account's default
signature into its ``HTMLBody``.
"""

from types import SimpleNamespace

from clariot.adapters.outlook import OutlookClient, _merge_signature

# Shaped like what the templates render: a fragment, not a whole document.
REPORT = f'<div style="font-family: Calibri;"><p>Estimado {OutlookClient.UNTOUCHED_MARKER}</p></div>'

# Shaped like what Outlook really returns: a whole document, with the styles and
# the inline logo the technician wants kept exactly as they are.
SIGNATURE = (
    "<html><head><style>p {font-family: Calibri;}</style></head>"
    '<body lang="ES"><div class="WordSection1">'
    '<p>Nicolas Lira<br><img src="cid:logo.png"></p>'
    "</div></body></html>"
)


class FakeAttachments:
    """Just enough of the Attachments collection for the writing paths."""

    def __init__(self) -> None:
        self.paths: list[str] = []

    @property
    def Count(self) -> int:
        return len(self.paths)

    def Add(self, path: str) -> None:
        self.paths.append(path)

    def Item(self, index: int):
        return SimpleNamespace(Delete=lambda: self.paths.pop(index - 1))


class FakeMail:
    """A mail item that signs itself when its Inspector is asked for."""

    EntryID = "DRAFT-1"

    def __init__(self, log: list[str], signature: str, inspector_error: str) -> None:
        self._signature = signature
        self._inspector_error = inspector_error
        self._log = log
        self._account = None
        self.Subject = ""
        self.HTMLBody = ""
        self.To = ""
        self.CC = ""
        self.Sent = False
        self.Attachments = FakeAttachments()
        self.saves = 0

    @property
    def GetInspector(self):
        """Reading this is what makes Outlook insert the signature."""
        self._log.append("inspector")
        if self._inspector_error:
            raise RuntimeError(self._inspector_error)
        self.HTMLBody = self._signature
        return object()

    @property
    def SendUsingAccount(self):
        return self._account

    @SendUsingAccount.setter
    def SendUsingAccount(self, account) -> None:
        self._log.append("account")
        self._account = account

    def Save(self) -> None:
        self.saves += 1


class FakeAccounts:
    def __init__(self, smtp: str) -> None:
        self._accounts = [SimpleNamespace(SmtpAddress=smtp)] if smtp else []

    @property
    def Count(self) -> int:
        return len(self._accounts)

    def Item(self, index: int):
        return self._accounts[index - 1]


def client(
    signature: str = SIGNATURE,
    *,
    use_signature: bool = True,
    inspector_error: str = "",
    account_smtp: str = "",
    existing_draft: FakeMail | None = None,
):
    """An OutlookClient wired to a fake Outlook, plus its item log.

    Returns the adapter, the list of items it created (the draft first, the
    throwaway probe that carries the signature second) and the ordered log of
    what it did to them.
    """
    log: list[str] = []
    created: list[FakeMail] = []

    def create_item(_kind: int) -> FakeMail:
        log.append("create")
        mail = FakeMail(log, signature, inspector_error)
        created.append(mail)
        return mail

    adapter = object.__new__(OutlookClient)
    adapter._app = SimpleNamespace(CreateItem=create_item)
    adapter._namespace = SimpleNamespace(
        Accounts=FakeAccounts(account_smtp),
        GetItemFromID=lambda entry_id: existing_draft,
    )
    adapter._use_signature = use_signature
    adapter._signature_cache = None
    return adapter, created, log


def draft_body(adapter, created, **kwargs) -> str:
    adapter.create_draft(
        subject="[ALERTA TECNICA] Reporte de Estado - Bomba 3",
        html_body=REPORT,
        to=(),
        cc=("porellana@emeltec.cl",),
        attachments=(),
        **kwargs,
    )
    return created[0].HTMLBody


def test_the_signature_ends_up_in_the_draft_below_the_report():
    """The whole point: the draft arrives signed, report first."""
    adapter, created, _ = client()

    body = draft_body(adapter, created)

    assert "Nicolas Lira" in body
    assert 'src="cid:logo.png"' in body
    assert body.index(OutlookClient.UNTOUCHED_MARKER) < body.index("Nicolas Lira")


def test_the_merged_body_is_a_single_document():
    """Two <html> roots render unpredictably, so the report is spliced in."""
    adapter, created, _ = client()

    body = draft_body(adapter, created)

    assert body.count("<html>") == 1
    assert body.startswith("<html><head>")
    # The signature's own <head> styles survive, above the report.
    assert body.index("<style>") < body.index(OutlookClient.UNTOUCHED_MARKER)


def test_the_merged_body_still_carries_the_untouched_marker():
    """``draft_state`` decides a draft is reusable by finding this marker.

    Losing it while merging would make every draft look edited by a human, and
    the second alarm would open a separate draft instead of growing this one.
    """
    adapter, created, _ = client()

    body = draft_body(adapter, created)

    assert OutlookClient.UNTOUCHED_MARKER in body


def test_an_account_without_a_signature_leaves_the_body_byte_identical():
    """Not having a signature configured is ordinary, not a failure."""
    adapter, created, _ = client(signature="")

    assert draft_body(adapter, created) == REPORT


def test_a_signature_that_is_only_a_fragment_is_still_merged():
    """Nothing guarantees a <body> tag, and the draft must still be signed."""
    adapter, created, _ = client(signature="<p>Nicolas Lira</p>")

    body = draft_body(adapter, created)

    assert body == REPORT + "<p>Nicolas Lira</p>"
    assert body.index(OutlookClient.UNTOUCHED_MARKER) < body.index("Nicolas Lira")


def test_the_setting_switches_the_signature_off():
    """use_outlook_signature: false keeps exactly the old behaviour."""
    adapter, created, log = client(use_signature=False)

    assert draft_body(adapter, created) == REPORT
    # Not even asked for: no throwaway item, no Inspector.
    assert log == ["create"]


def test_an_inspector_that_refuses_does_not_break_draft_creation():
    """A signature is a nicety; the draft is the job. It must never block it."""
    adapter, created, _ = client(inspector_error="Outlook refused the Inspector")

    assert draft_body(adapter, created) == REPORT
    assert created[0].saves == 1


def test_the_throwaway_item_is_never_saved():
    """Saving it would drop an empty message into the Drafts folder."""
    adapter, created, _ = client()

    draft_body(adapter, created)

    probe = created[1]
    assert probe.saves == 0


def test_the_sending_account_is_chosen_before_the_signature_is_read():
    """Outlook offers the signature of the account the message is sent from."""
    adapter, created, log = client(account_smtp="reportes@emeltec.cl")

    draft_body(adapter, created, sender_account="reportes@emeltec.cl")

    assert log.index("account") < log.index("inspector")


def test_the_signature_is_read_once_per_run():
    """One throwaway item and one Inspector, however many drafts are written."""
    adapter, created, log = client()

    draft_body(adapter, created)
    adapter.create_draft(
        subject="otra alarma", html_body=REPORT, to=(), cc=(), attachments=()
    )

    assert log.count("inspector") == 1
    assert created[-1].HTMLBody == created[0].HTMLBody


def test_growing_an_existing_draft_keeps_it_signed():
    """``update_draft`` rewrites the body wholesale, signature included."""
    existing = FakeMail([], SIGNATURE, "")
    existing.Parent = SimpleNamespace(Name="Por enviar")
    adapter, _, _ = client(existing_draft=existing)

    adapter.update_draft("DRAFT-1", subject="dos alarmas", html_body=REPORT)

    assert "Nicolas Lira" in existing.HTMLBody
    assert OutlookClient.UNTOUCHED_MARKER in existing.HTMLBody


def test_merging_an_empty_signature_changes_nothing():
    """The unit behind every "no signature configured" path."""
    assert _merge_signature(REPORT, "") == REPORT
    assert _merge_signature(REPORT, "   \n") == REPORT


def test_the_body_tag_is_found_whatever_its_case_or_attributes():
    """Outlook writes <BODY> and adds attributes; a literal match would miss."""
    merged = _merge_signature("<p>informe</p>", '<HTML><BODY LANG="ES">firma</BODY></HTML>')

    assert merged == '<HTML><BODY LANG="ES"><p>informe</p>firma</BODY></HTML>'
