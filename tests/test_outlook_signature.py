"""The draft is signed by Outlook itself, so nobody pastes a signature by hand.

No real Outlook is involved. ``OutlookClient.__init__`` dispatches COM, which
only exists on a Windows machine running classic Outlook, so the instance is
built without it and handed a fake application that behaves the way Outlook
does: touching ``GetInspector`` on a mail item drops the account's default
signature into its ``HTMLBody``.
"""

import base64
from pathlib import Path
from types import SimpleNamespace

import pytest

import clariot.adapters.outlook as outlook_module
from clariot.adapters.outlook import (
    PROP_ATTACH_CONTENT_ID,
    PROP_ATTACH_HIDDEN,
    OutlookClient,
    _merge_signature,
)

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


@pytest.fixture(autouse=True)
def _no_registry_default(monkeypatch):
    """The default signature name comes from the registry; tests set their own."""
    monkeypatch.setattr(outlook_module, "_default_signature_name", lambda: "")


class FakeAttachment:
    """An attachment that knows its Content-ID the way a real inline image does."""

    def __init__(
        self, owner, *, path: str = "", file_name: str = "", cid: str = "", save_error: str = ""
    ) -> None:
        self.path = path
        self._save_error = save_error
        self.FileName = file_name or Path(path).name
        self.props: dict[str, object] = {PROP_ATTACH_CONTENT_ID: cid} if cid else {}
        self._owner = owner
        self.PropertyAccessor = SimpleNamespace(GetProperty=self._get, SetProperty=self._set)

    def _get(self, name: str):
        # Real Outlook raises for a property the attachment does not have.
        if name not in self.props:
            raise KeyError(name)
        return self.props[name]

    def _set(self, name: str, value) -> None:
        self.props[name] = value

    def SaveAsFile(self, path: str) -> None:
        if self._save_error:
            raise RuntimeError(self._save_error)
        Path(path).write_bytes(b"IMAGE")

    def Delete(self) -> None:
        self._owner.items.remove(self)


class FakeAttachments:
    """Just enough of the Attachments collection for the writing paths."""

    def __init__(self, add_error: str = "") -> None:
        self.items: list[FakeAttachment] = []
        self._add_error = add_error

    @property
    def paths(self) -> list[str]:
        return [item.path for item in self.items]

    @property
    def Count(self) -> int:
        return len(self.items)

    def Add(self, path: str) -> FakeAttachment:
        if self._add_error:
            raise RuntimeError(self._add_error)
        attachment = FakeAttachment(self, path=path)
        self.items.append(attachment)
        return attachment

    def Item(self, index: int) -> FakeAttachment:
        return self.items[index - 1]


class FakeMail:
    """A mail item that signs itself when its Inspector is asked for."""

    EntryID = "DRAFT-1"

    def __init__(
        self,
        log: list[str],
        signature: str,
        inspector_error: str,
        images: dict[str, str] | None = None,
        add_error: str = "",
        save_error: str = "",
    ) -> None:
        self._save_error = save_error
        self._signature = signature
        self._inspector_error = inspector_error
        self._images = images or {}
        self._log = log
        self._account = None
        self.Subject = ""
        self.HTMLBody = ""
        self.To = ""
        self.CC = ""
        self.Sent = False
        self.Attachments = FakeAttachments(add_error)
        self.saves = 0

    @property
    def GetInspector(self):
        """Reading this is what makes Outlook insert the signature.

        A signature's logo arrives as a hidden attachment of this same item,
        referenced from the HTML by its Content-ID.
        """
        self._log.append("inspector")
        if self._inspector_error:
            raise RuntimeError(self._inspector_error)
        self.HTMLBody = self._signature
        for cid, file_name in self._images.items():
            self.Attachments.items.append(
                FakeAttachment(
                    self.Attachments, file_name=file_name, cid=cid, save_error=self._save_error
                )
            )
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
    images: dict[str, str] | None = None,
    add_error: str = "",
    save_error: str = "",
    signatures_dir: Path | None = None,
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
        mail = FakeMail(log, signature, inspector_error, images, add_error, save_error)
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
    # Never the real folder of whoever runs the tests.
    adapter._signatures_dir = signatures_dir or Path("no-signatures-here")
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


# The signature as a real classic Outlook returned it: no text at all, a single
# picture referenced by Content-ID, and the picture itself a hidden attachment.
REAL_CID = "image001.jpg@01DD5FC8.9FF38EE0"
IMAGE_ONLY_SIGNATURE = (
    "<html><head><style>p.MsoNormal {margin: 0;}</style></head>"
    "<body lang=ES-CL><div class=WordSection1><p class=MsoNormal><o:p>&nbsp;</o:p></p>"
    '<p class=MsoNormal><a name="_MailAutoSig"><span>'
    f'<img width=615 height=133 id="Imagen_x0020_1" src="cid:{REAL_CID}">'
    "</span></a></p></div></body></html>"
)


def test_a_picture_only_signature_is_placed_below_the_report():
    adapter, created, _ = client(
        signature=IMAGE_ONLY_SIGNATURE, images={REAL_CID: "image001.jpg"}
    )

    body = draft_body(adapter, created)

    assert body.index(OutlookClient.UNTOUCHED_MARKER) < body.index(f"cid:{REAL_CID}")


def test_the_signature_image_travels_with_the_draft():
    """The regression: the HTML pointed at an image the draft did not carry.

    A logo is referenced as ``cid:`` and lives on the item Outlook signed, so
    copying the HTML alone left a broken picture, which for a signature that is
    only a picture means no visible signature at all.
    """
    adapter, created, _ = client(
        signature=IMAGE_ONLY_SIGNATURE, images={REAL_CID: "image001.jpg"}
    )

    draft_body(adapter, created)

    (image,) = created[0].Attachments.items
    assert image.props[PROP_ATTACH_CONTENT_ID] == REAL_CID
    assert image.props[PROP_ATTACH_HIDDEN] is True
    assert Path(image.path).read_bytes() == b"IMAGE"


def test_an_attachment_the_signature_does_not_reference_is_not_copied():
    adapter, created, _ = client(
        signature=IMAGE_ONLY_SIGNATURE, images={"otra.png@01": "otra.png"}
    )

    draft_body(adapter, created)

    assert created[0].Attachments.items == []


def test_a_signature_without_images_attaches_nothing():
    adapter, created, _ = client()

    draft_body(adapter, created)

    assert created[0].Attachments.items == []


def test_growing_a_draft_keeps_the_signature_image():
    """``update_draft`` clears every attachment, the hidden image included."""
    existing = FakeMail([], IMAGE_ONLY_SIGNATURE, "")
    existing.Parent = SimpleNamespace(Name="Por enviar")
    existing.Attachments.Add("C:/old/previous-alarm.pdf")
    adapter, _, _ = client(
        signature=IMAGE_ONLY_SIGNATURE,
        images={REAL_CID: "image001.jpg"},
        existing_draft=existing,
    )

    adapter.update_draft("DRAFT-1", subject="dos alarmas", html_body=REPORT)

    (image,) = existing.Attachments.items
    assert image.props[PROP_ATTACH_CONTENT_ID] == REAL_CID
    assert f"cid:{REAL_CID}" in existing.HTMLBody


def test_the_signature_image_is_read_once_per_run():
    adapter, created, _ = client(
        signature=IMAGE_ONLY_SIGNATURE, images={REAL_CID: "image001.jpg"}
    )

    draft_body(adapter, created)
    adapter.create_draft(subject="otra", html_body=REPORT, to=(), cc=(), attachments=())

    drafts = [mail for mail in created if mail.saves]
    assert len(drafts) == 2
    assert all(len(mail.Attachments.items) == 1 for mail in drafts)


def test_a_signature_image_that_cannot_be_attached_does_not_block_the_draft():
    """A missing picture is a flaw; a missing draft is the failure."""
    adapter, created, _ = client(
        signature=IMAGE_ONLY_SIGNATURE,
        images={REAL_CID: "image001.jpg"},
        add_error="Outlook refused the attachment",
    )

    body = draft_body(adapter, created)

    assert created[0].saves == 1
    assert OutlookClient.UNTOUCHED_MARKER in body


# --- a picture Outlook cannot hand over ------------------------------------

PICTURE = b"PICTURE-BYTES"


def signature_file(folder: Path, name: str, payload: bytes = PICTURE, kind: str = "jpeg") -> Path:
    """A signature file the way a data-URI signature is stored: no _archivos folder."""
    path = folder / f"{name}.htm"
    path.write_bytes(
        b'<html><body><p><img width=615 src="data:image/'
        + kind.encode()
        + b";base64,"
        + base64.b64encode(payload)
        + b'"></p></body></html>'
    )
    return path


def test_a_picture_outlook_cannot_save_is_read_from_the_signature_file(tmp_path):
    """The case on the technician's PC.

    The signature's logo is a data URI inside the .htm. Outlook lists it as an
    attachment of the throwaway item but refuses to save it ("the data source is
    missing"), so the bytes have to come from the file.
    """
    signature_file(tmp_path, "pie de firma")
    adapter, created, _ = client(
        signature=IMAGE_ONLY_SIGNATURE,
        images={REAL_CID: "image001.jpg"},
        save_error="No se pueden guardar los datos adjuntos. Falta el origen de los datos.",
        signatures_dir=tmp_path,
    )

    draft_body(adapter, created)

    (image,) = created[0].Attachments.items
    assert image.props[PROP_ATTACH_CONTENT_ID] == REAL_CID
    assert image.props[PROP_ATTACH_HIDDEN] is True
    assert Path(image.path).read_bytes() == PICTURE
    assert image.path.endswith(".jpg")


def test_a_probe_with_no_attachments_at_all_still_gets_its_picture(tmp_path):
    """The HTML names a cid, but Outlook lists nothing: the file is the only source."""
    signature_file(tmp_path, "pie de firma")
    adapter, created, _ = client(signature=IMAGE_ONLY_SIGNATURE, signatures_dir=tmp_path)

    draft_body(adapter, created)

    (image,) = created[0].Attachments.items
    assert image.props[PROP_ATTACH_CONTENT_ID] == REAL_CID
    assert Path(image.path).read_bytes() == PICTURE


def test_without_a_signature_file_the_draft_is_still_signed_and_created(tmp_path, caplog):
    """No recoverable picture is a flaw in the signature, never a failed draft."""
    adapter, created, _ = client(signature=IMAGE_ONLY_SIGNATURE, signatures_dir=tmp_path)

    with caplog.at_level("INFO"):
        body = draft_body(adapter, created)

    assert created[0].saves == 1
    assert created[0].Attachments.items == []
    assert f"cid:{REAL_CID}" in body
    assert "Could not read the default Outlook signature" not in caplog.text


def test_several_signature_files_and_no_default_pick_nothing(tmp_path):
    """Guessing could sign a report with somebody else's signature."""
    signature_file(tmp_path, "firma uno")
    signature_file(tmp_path, "firma dos", payload=b"OTHER")
    adapter, created, _ = client(signature=IMAGE_ONLY_SIGNATURE, signatures_dir=tmp_path)

    draft_body(adapter, created)

    assert created[0].Attachments.items == []


def test_the_default_signature_name_chooses_between_several_files(tmp_path, monkeypatch):
    signature_file(tmp_path, "firma uno", payload=b"FIRST")
    signature_file(tmp_path, "firma dos", payload=b"SECOND")
    monkeypatch.setattr(outlook_module, "_default_signature_name", lambda: "Firma Dos")
    adapter, created, _ = client(signature=IMAGE_ONLY_SIGNATURE, signatures_dir=tmp_path)

    draft_body(adapter, created)

    (image,) = created[0].Attachments.items
    assert Path(image.path).read_bytes() == b"SECOND"


def test_pictures_are_read_in_document_order_and_tolerate_wrapped_base64(tmp_path):
    wrapped = base64.encodebytes(b"A" * 200)  # encodebytes inserts newlines
    (tmp_path / "firma.htm").write_bytes(
        b'<img src="data:image/png;base64,' + wrapped + b'">'
        b'<img src="data:image/jpeg;base64,' + base64.b64encode(b"B") + b'">'
    )

    images = outlook_module._embedded_signature_images(tmp_path, "")

    assert images == [(b"A" * 200, "png"), (b"B", "jpg")]


def test_a_signature_file_without_pictures_yields_none(tmp_path):
    (tmp_path / "firma.htm").write_bytes(b"<html><body><p>solo texto</p></body></html>")

    assert outlook_module._embedded_signature_images(tmp_path, "") == []


def test_a_folder_that_does_not_exist_yields_none(tmp_path):
    assert outlook_module._embedded_signature_images(tmp_path / "nope", "") == []
