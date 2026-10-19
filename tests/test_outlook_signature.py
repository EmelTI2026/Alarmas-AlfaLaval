"""The draft is signed from the signature file the user already has.

Classic Outlook keeps each signature as an .htm file in the user's profile, so
the program reads it from there. That works whether or not a default signature
is configured, and it does not depend on Outlook inserting anything itself.

No real Outlook is involved. ``OutlookClient.__init__`` dispatches COM, which
only exists on a Windows machine running classic Outlook, so the instance is
built without it and handed a fake application whose mail items record what was
done to them and in which order.
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
REPORT = f'<div style="font-family: Calibri;"><p>Estimado [NOMBRE] {OutlookClient.UNTOUCHED_MARKER}</p></div>'

PICTURE = b"PICTURE-BYTES"
DATA_URI = "data:image/jpeg;base64," + base64.b64encode(PICTURE).decode()


def signature_html(image_src: str, text: str = "Firma") -> str:
    """A signature document shaped like the ones Outlook writes."""
    return (
        '<html><head><meta http-equiv=Content-Type content="text/html; charset=windows-1252">'
        "<style>p.MsoNormal {margin: 0;}</style></head>"
        '<body lang=ES-CL><div class=WordSection1><p class=MsoNormal>'
        f'{text}<br><img width=615 height=133 src="{image_src}"></p></div></body></html>'
    )


def write_signature(folder: Path, name: str, html: str, encoding: str = "cp1252") -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"{name}.htm"
    path.write_bytes(html.encode(encoding))
    return path


@pytest.fixture(autouse=True)
def _no_registry_default(monkeypatch):
    """The default signature name comes from the registry; tests set their own."""
    monkeypatch.setattr(outlook_module, "_default_signature_name", lambda: "")


class FakeAttachment:
    def __init__(self, owner, path: str) -> None:
        self.path = path
        self.props: dict[str, object] = {}
        self._owner = owner
        self.PropertyAccessor = SimpleNamespace(SetProperty=self._set)

    def _set(self, name: str, value) -> None:
        self.props[name] = value

    def Delete(self) -> None:
        self._owner.items.remove(self)
        self._owner.ops.append("delete")


class FakeAttachments:
    """Just enough of the Attachments collection for the writing paths."""

    def __init__(self, ops: list[str], add_error: str = "") -> None:
        self.items: list[FakeAttachment] = []
        self.ops = ops
        self._add_error = add_error

    @property
    def Count(self) -> int:
        return len(self.items)

    def Add(self, path: str) -> FakeAttachment:
        if self._add_error:
            raise RuntimeError(self._add_error)
        self.ops.append("attach")
        attachment = FakeAttachment(self, path)
        self.items.append(attachment)
        return attachment

    def Item(self, index: int) -> FakeAttachment:
        return self.items[index - 1]


class FakeMail:
    """A mail item that records what is done to it, in order."""

    EntryID = "DRAFT-1"

    def __init__(self, add_error: str = "") -> None:
        self.ops: list[str] = []
        self._html = ""
        self.Subject = ""
        self.To = ""
        self.CC = ""
        self.Sent = False
        self.Attachments = FakeAttachments(self.ops, add_error)
        self.saves = 0

    @property
    def HTMLBody(self) -> str:
        return self._html

    @HTMLBody.setter
    def HTMLBody(self, value: str) -> None:
        self._html = value
        self.ops.append("body")

    def Save(self) -> None:
        self.saves += 1


def client(
    signatures_dir: Path | None = None,
    *,
    use_signature: bool = True,
    account: str = "",
    existing_draft: FakeMail | None = None,
    add_error: str = "",
):
    """An OutlookClient wired to a fake Outlook, plus the items it created."""
    created: list[FakeMail] = []

    def create_item(_kind: int) -> FakeMail:
        mail = FakeMail(add_error)
        created.append(mail)
        return mail

    adapter = object.__new__(OutlookClient)
    adapter._app = SimpleNamespace(CreateItem=create_item)
    adapter._namespace = SimpleNamespace(GetItemFromID=lambda entry_id: existing_draft)
    adapter._use_signature = use_signature
    adapter._signature_cache = None
    adapter._signature_account = account
    # Never the real folder of whoever runs the tests.
    adapter._signatures_dir = signatures_dir or Path("no-signatures-here")
    return adapter, created


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


# --- the signature ends up in the draft -------------------------------------


def test_the_draft_is_signed_with_the_users_signature_file(tmp_path):
    """The whole point: the draft arrives signed, report first."""
    write_signature(tmp_path, "pie de firma", signature_html(DATA_URI, "Nicolas Lira"))
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    assert "Nicolas Lira" in body
    assert body.index(OutlookClient.UNTOUCHED_MARKER) < body.index("Nicolas Lira")


def test_a_picture_embedded_in_the_file_is_left_exactly_as_it_is(tmp_path):
    """The case on the technician's PC.

    Their logo is a base64 data URI inside the .htm, with no folder beside it.
    Outlook renders it in the draft as it does in a message started by hand, so
    it needs no attachment and no rewriting.
    """
    write_signature(tmp_path, "pie de firma", signature_html(DATA_URI))
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    assert DATA_URI in body
    assert created[0].Attachments.items == []


def test_the_merged_body_is_a_single_document(tmp_path):
    """Two <html> roots render unpredictably, so the report is spliced in."""
    write_signature(tmp_path, "firma", signature_html(DATA_URI))
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    assert body.lower().count("<html") == 1
    assert body.startswith("<html><head>")
    # The signature's own <head> styles survive, above the report.
    assert body.index("<style>") < body.index(OutlookClient.UNTOUCHED_MARKER)


def test_the_merged_body_still_carries_the_untouched_marker(tmp_path):
    """``draft_state`` decides a draft is reusable by finding this marker.

    Losing it while merging would make every draft look edited by a human, and
    the second alarm would open a separate draft instead of growing this one.
    """
    write_signature(tmp_path, "firma", signature_html(DATA_URI))
    adapter, created = client(tmp_path)

    assert OutlookClient.UNTOUCHED_MARKER in draft_body(adapter, created)


def test_no_signature_leaves_the_body_byte_identical(tmp_path):
    """Not having a signature is ordinary, not a failure."""
    adapter, created = client(tmp_path / "does-not-exist")

    assert draft_body(adapter, created) == REPORT
    assert created[0].saves == 1


def test_an_empty_signature_folder_leaves_the_body_byte_identical(tmp_path):
    adapter, created = client(tmp_path)

    assert draft_body(adapter, created) == REPORT


def test_the_setting_switches_the_signature_off(tmp_path):
    """use_outlook_signature: false keeps exactly the old behaviour."""
    write_signature(tmp_path, "firma", signature_html(DATA_URI))
    adapter, created = client(tmp_path, use_signature=False)

    assert draft_body(adapter, created) == REPORT


def test_a_signature_that_is_only_a_fragment_is_still_merged(tmp_path):
    """Nothing guarantees a <body> tag, and the draft must still be signed."""
    write_signature(tmp_path, "firma", "<p>Nicolas Lira</p>")
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    assert body == REPORT + "<p>Nicolas Lira</p>"


# --- which signature ---------------------------------------------------------


def test_a_lone_signature_is_used_when_no_default_is_set(tmp_path):
    """Plenty of accounts have no default; one file is unambiguous."""
    write_signature(tmp_path, "pie de firma (jlillo@emeltec.cl)", signature_html(DATA_URI, "UNICA"))
    adapter, created = client(tmp_path)

    assert "UNICA" in draft_body(adapter, created)


def test_the_default_signature_wins_among_several(tmp_path, monkeypatch):
    write_signature(tmp_path, "firma uno", signature_html(DATA_URI, "PRIMERA"))
    write_signature(tmp_path, "firma dos", signature_html(DATA_URI, "SEGUNDA"))
    monkeypatch.setattr(outlook_module, "_default_signature_name", lambda: "Firma Dos")
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    assert "SEGUNDA" in body
    assert "PRIMERA" not in body


def test_the_sending_account_picks_its_own_signature_among_several(tmp_path):
    write_signature(tmp_path, "firma (otro@emeltec.cl)", signature_html(DATA_URI, "OTRO"))
    write_signature(tmp_path, "firma (JLILLO@emeltec.cl)", signature_html(DATA_URI, "MIA"))
    adapter, created = client(tmp_path, account="jlillo@emeltec.cl")

    body = draft_body(adapter, created)

    assert "MIA" in body
    assert "OTRO" not in body


def test_several_signatures_and_no_way_to_choose_leave_the_draft_unsigned(tmp_path, caplog):
    """Guessing could sign a report with somebody else's signature."""
    write_signature(tmp_path, "firma uno", signature_html(DATA_URI, "PRIMERA"))
    write_signature(tmp_path, "firma dos", signature_html(DATA_URI, "SEGUNDA"))
    adapter, created = client(tmp_path)

    with caplog.at_level("WARNING"):
        body = draft_body(adapter, created)

    assert body == REPORT
    assert "none is the default" in caplog.text


def test_the_signature_is_read_once_per_run(tmp_path):
    path = write_signature(tmp_path, "firma", signature_html(DATA_URI, "FIRMA-DE-HOY"))
    adapter, created = client(tmp_path)

    draft_body(adapter, created)
    path.unlink()
    adapter.create_draft(subject="otra", html_body=REPORT, to=(), cc=(), attachments=())

    assert "FIRMA-DE-HOY" in created[1].HTMLBody


# --- reading the file ---------------------------------------------------------


def test_the_declared_windows_encoding_is_honoured(tmp_path):
    write_signature(tmp_path, "firma", signature_html(DATA_URI, "Atención Ñandú"), "cp1252")
    adapter, created = client(tmp_path)

    assert "Atención Ñandú" in draft_body(adapter, created)


def test_a_utf8_signature_is_read_too(tmp_path):
    html = signature_html(DATA_URI, "Atención Ñandú").replace("windows-1252", "utf-8")
    write_signature(tmp_path, "firma", html, "utf-8")
    adapter, created = client(tmp_path)

    assert "Atención Ñandú" in draft_body(adapter, created)


def test_an_unreadable_signature_never_blocks_the_draft(tmp_path, monkeypatch, caplog):
    """A signature is a nicety; the draft is the job."""

    def broken(*_args, **_kwargs):
        raise OSError("the file is locked")

    monkeypatch.setattr(outlook_module, "_load_signature", broken)
    adapter, created = client(tmp_path)

    with caplog.at_level("WARNING"):
        body = draft_body(adapter, created)

    assert body == REPORT
    assert created[0].saves == 1
    assert "Could not read the signature file" in caplog.text


# --- pictures kept beside the file --------------------------------------------


def classic_signature(tmp_path: Path, src: str = "firma_archivos/image001.png") -> Path:
    """A classic signature: the .htm plus a folder holding its picture."""
    (tmp_path / "firma_archivos").mkdir()
    picture = tmp_path / "firma_archivos" / "image001.png"
    picture.write_bytes(PICTURE)
    write_signature(tmp_path, "firma", signature_html(src))
    return picture


def test_a_picture_kept_in_a_folder_is_attached_and_referenced_by_cid(tmp_path):
    """A relative path would show up broken in a draft, so it travels as inline."""
    picture = classic_signature(tmp_path)
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    (image,) = created[0].Attachments.items
    assert Path(image.path) == picture.resolve()
    assert image.props[PROP_ATTACH_CONTENT_ID] == "clariot-firma-1"
    assert image.props[PROP_ATTACH_HIDDEN] is True
    assert 'src="cid:clariot-firma-1"' in body
    assert "firma_archivos" not in body


def test_the_pictures_are_attached_before_the_body_is_written(tmp_path):
    """Outlook links a cid: reference to an attachment that already exists.

    With the body written first, the picture showed up as a plain file in the
    attachment list instead of inside the body as the signature.
    """
    classic_signature(tmp_path)
    adapter, created = client(tmp_path)

    draft_body(adapter, created)

    ops = created[0].ops
    assert ops.index("attach") < ops.index("body")


def test_a_picture_that_does_not_exist_is_left_alone(tmp_path):
    write_signature(tmp_path, "firma", signature_html("firma_archivos/missing.png"))
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    assert "firma_archivos/missing.png" in body
    assert created[0].Attachments.items == []


def test_a_path_that_climbs_out_of_the_signature_folder_is_never_attached(tmp_path):
    """A signature file must not be able to pull in an arbitrary file."""
    folder = tmp_path / "Signatures"
    folder.mkdir()
    (tmp_path / "secret.png").write_bytes(b"SECRET")
    write_signature(folder, "firma", signature_html("../secret.png"))
    adapter, created = client(folder)

    draft_body(adapter, created)

    assert created[0].Attachments.items == []


@pytest.mark.parametrize(
    "src",
    [DATA_URI, "cid:logo.png@01", "https://example.com/logo.png", "file:///C:/logo.png"],
)
def test_only_relative_pictures_are_rewritten(tmp_path, src):
    write_signature(tmp_path, "firma", signature_html(src))
    adapter, created = client(tmp_path)

    body = draft_body(adapter, created)

    assert src in body
    assert created[0].Attachments.items == []


def test_a_picture_that_cannot_be_attached_does_not_block_the_draft(tmp_path):
    """A missing picture is a flaw; a missing draft is the failure."""
    classic_signature(tmp_path)
    adapter, created = client(tmp_path, add_error="Outlook refused the attachment")

    body = draft_body(adapter, created)

    assert created[0].saves == 1
    assert OutlookClient.UNTOUCHED_MARKER in body


# --- growing an existing draft --------------------------------------------------


def existing_draft(*, with_old_pdf: bool = True) -> FakeMail:
    existing = FakeMail()
    existing.Parent = SimpleNamespace(Name="Por enviar")
    if with_old_pdf:
        existing.Attachments.Add("C:/old/previous-alarm.pdf")
    existing.ops.clear()
    return existing


def test_growing_an_existing_draft_keeps_it_signed(tmp_path):
    """``update_draft`` rewrites the body wholesale, signature included."""
    write_signature(tmp_path, "firma", signature_html(DATA_URI, "Nicolas Lira"))
    existing = existing_draft()
    adapter, _ = client(tmp_path, existing_draft=existing)

    adapter.update_draft("DRAFT-1", subject="dos alarmas", html_body=REPORT)

    assert "Nicolas Lira" in existing.HTMLBody
    assert OutlookClient.UNTOUCHED_MARKER in existing.HTMLBody


def test_growing_a_draft_puts_the_pictures_back_before_the_body(tmp_path):
    """``update_draft`` clears every attachment, the signature's pictures included."""
    classic_signature(tmp_path)
    existing = existing_draft()
    adapter, _ = client(tmp_path, existing_draft=existing)

    adapter.update_draft("DRAFT-1", subject="dos alarmas", html_body=REPORT)

    assert existing.ops == ["delete", "attach", "body"]
    (image,) = existing.Attachments.items
    assert image.props[PROP_ATTACH_CONTENT_ID] == "clariot-firma-1"
    assert 'src="cid:clariot-firma-1"' in existing.HTMLBody


# --- the merge itself -----------------------------------------------------------


def test_merging_an_empty_signature_changes_nothing():
    """The unit behind every "no signature" path."""
    assert _merge_signature(REPORT, "") == REPORT
    assert _merge_signature(REPORT, "   \n") == REPORT


def test_the_body_tag_is_found_whatever_its_case_or_attributes():
    """Outlook writes <BODY> and adds attributes; a literal match would miss."""
    merged = _merge_signature("<p>informe</p>", '<HTML><BODY LANG="ES">firma</BODY></HTML>')

    assert merged == '<HTML><BODY LANG="ES"><p>informe</p>firma</BODY></HTML>'
