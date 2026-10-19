"""Outlook desktop integration through COM (pywin32).

Deployment constraints worth knowing before touching this file:

* Only *classic* Outlook for Windows exposes a COM automation surface. The new
  Outlook client does not, and no amount of patching here will change that.
* COM requires an interactive desktop session. A Scheduled Task configured with
  "Run whether user is logged on or not" fails under session 0 isolation. Use an
  "on logon" trigger with repetition instead, on a machine that stays signed in.
"""

from __future__ import annotations

import base64
import logging
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from ..textutils import sanitize_filename

logger = logging.getLogger(__name__)

OL_FOLDER_INBOX = 6
OL_FOLDER_DRAFTS = 16
OL_MAIL_ITEM = 0
OL_CLASS_MAIL = 43

# PR_INTERNET_MESSAGE_ID_W: stable across folder moves, unlike EntryID.
PROP_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
# PR_ATTACH_CONTENT_ID_W: present on inline images such as signature logos.
PROP_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
# PR_ATTACHMENT_HIDDEN: keeps an inline image out of the attachment list.
PROP_ATTACH_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


class OutlookError(RuntimeError):
    """Raised when Outlook is unreachable or a required folder is missing."""


def classic_outlook_registered() -> bool:
    """True when the classic Outlook COM server is registered on this machine.

    The new Outlook (``olk.exe``, a Store package) exposes no COM server, so this
    check separates "Outlook is closed" from "the required client is absent".
    """
    try:
        import winreg
    except ImportError:  # pragma: no cover - Windows-only
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Outlook.Application\CLSID"):
            return True
    except OSError:
        return False


NEW_OUTLOOK_HINT = (
    "Classic Outlook for Windows is not installed on this machine: the COM "
    "server 'Outlook.Application' is not registered. If the running client is "
    "olk.exe, that is the new Outlook, which exposes no local automation "
    "interface at all. This integration needs classic Outlook (the desktop app "
    "shipped with Microsoft 365 Apps)."
)


@dataclass(frozen=True)
class MailSummary:
    """One unread message as plain data, read before anything is mutated."""
    entry_id: str
    message_key: str
    subject: str
    received: str
    sender: str
    folder: str = ""
    """Path of the folder the message was found in, e.g. "1_Auditoria_Clariot/Nestle"."""


def _recipient_count(value: str) -> int:
    """How many entries an Outlook recipient string holds.

    Outlook writes ``;`` between entries, so that is the separator. A comma only
    splits a piece that already holds two addresses, which is how a line typed by
    hand can look before Outlook normalises it. Splitting on every comma would be
    wrong: a copy resolved against the address book reads back as
    ``"Orellana, Pablo"``, and counting that as two recipients is exactly the
    miscount this function exists to avoid.
    """
    entries = 0
    for part in value.split(";"):
        part = part.strip()
        if not part:
            continue
        entries += max(part.count("@"), 1)
    return entries


_BODY_OPEN_TAG = re.compile(r"<body\b[^>]*>", re.IGNORECASE)


def _merge_signature(html_body: str, signature: str) -> str:
    """Place our report inside the signature document, above the signature.

    Outlook hands the default signature back as a *complete* HTML document,
    styles and inline images included. Writing our body and the signature one
    after the other would give the draft two ``<html>`` roots, and Outlook
    renders that unpredictably. So the report is spliced in right after the
    signature's opening ``<body>`` tag instead: one well-formed document, the
    report first, the signature underneath it.

    An empty signature leaves the body byte-identical, which is what keeps the
    "no signature configured" case indistinguishable from the old behaviour.
    """
    if not signature.strip():
        return html_body

    match = _BODY_OPEN_TAG.search(signature)
    if match is None:
        # A fragment rather than a document. Concatenating is then correct, and
        # the ordering the technician expects still holds.
        return html_body + signature

    cut = match.end()
    return signature[:cut] + html_body + signature[cut:]


_CID_REFERENCE = re.compile(r"""cid:([^"'\s>)]+)""", re.IGNORECASE)

# A picture stored inside the signature file itself, as a base64 data URI.
_DATA_URI_IMAGE = re.compile(
    rb"""<img\b[^>]*?\bsrc=["']data:image/(?P<kind>[A-Za-z0-9.+-]+);base64,(?P<data>[^"']+)["']""",
    re.IGNORECASE | re.DOTALL,
)


def _default_signature_name() -> str:
    """Name Outlook keeps for the default new-message signature, or "".

    Read from the registry because the signature files themselves do not say
    which one is the default.
    """
    try:
        import winreg
    except ImportError:  # pragma: no cover - Windows-only
        return ""
    for version in ("16.0", "15.0", "14.0"):
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                rf"Software\Microsoft\Office\{version}\Common\MailSettings",
            ) as key:
                value, _ = winreg.QueryValueEx(key, "NewSignature")
        except OSError:
            continue
        if value:
            return str(value)
    return ""


def _pick_signature_file(folder: Path, preferred: str) -> Path | None:
    """The signature file to read pictures from, or None when it is ambiguous.

    The default signature wins when its name is known. Without it, a lone file is
    unambiguous; several files and no way to tell them apart means guessing, and
    a guess could sign a report with somebody else's signature.
    """
    if not folder.is_dir():
        return None
    files = sorted(folder.glob("*.htm"))
    for file in files:
        if preferred and file.stem.lower() == preferred.lower():
            return file
    return files[0] if len(files) == 1 else None


def _embedded_signature_images(folder: Path, preferred: str) -> list[tuple[bytes, str]]:
    """Pictures stored as data URIs inside the signature file, in document order.

    Outlook cannot hand these over as attachments (saving them fails with "the
    data source is missing"), but the bytes are right there in the file.
    """
    file = _pick_signature_file(folder, preferred)
    if file is None:
        return []
    images: list[tuple[bytes, str]] = []
    for match in _DATA_URI_IMAGE.finditer(file.read_bytes()):
        try:
            data = base64.b64decode(re.sub(rb"\s+", b"", match.group("data")))
        except ValueError:
            continue
        kind = match.group("kind").decode("ascii").lower().replace("jpeg", "jpg")
        images.append((data, kind))
    return images


def _safe_property(item, prop: str) -> str:
    try:
        return str(item.PropertyAccessor.GetProperty(prop) or "")
    except Exception:  # noqa: BLE001 - COM raises bare pywintypes.com_error
        return ""


class OutlookClient:
    """Thin wrapper over the Outlook object model.

    Kept deliberately small so the rest of the pipeline can be tested against a
    fake implementing the same handful of methods.
    """

    # Class-level default so an instance built without __init__ (as the tests do,
    # since __init__ dispatches COM) still has no images to attach.
    _signature_images: tuple[tuple[Path, str], ...] = ()
    # Where Outlook keeps signature files. None means the current user's folder;
    # the tests point it at a temporary directory.
    _signatures_dir: Path | None = None

    def __init__(self, *, use_signature: bool = True) -> None:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:  # pragma: no cover - Windows-only dependency
            raise OutlookError("pywin32 is required to talk to Outlook") from exc

        self._use_signature = use_signature
        # Read at most once per run, and only if a draft is actually written.
        self._signature_cache: str | None = None
        # The signature's hidden images, as (saved file, Content-ID) pairs.
        self._signature_images: tuple[tuple[Path, str], ...] = ()

        # Required when running from a thread or a scheduled task host.
        pythoncom.CoInitialize()
        try:
            self._app = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._app.GetNamespace("MAPI")
        except Exception as exc:  # noqa: BLE001
            if not classic_outlook_registered():
                raise OutlookError(NEW_OUTLOOK_HINT) from exc
            raise OutlookError(
                "Classic Outlook is installed but did not respond. Make sure it "
                f"is running and signed in. Underlying error: {exc}"
            ) from exc

    # ------------------------------------------------------------------ folders

    def _inbox(self):
        return self._namespace.GetDefaultFolder(OL_FOLDER_INBOX)

    def _find_folder(self, name: str):
        """Breadth-first, case-insensitive search under the Inbox.

        Accepts a nested path such as ``"Clientes/1_Auditoria_Clariot"``.
        """
        target = name.replace("\\", "/").strip("/")
        current = self._inbox()
        for part in target.split("/"):
            wanted = part.strip().lower()
            match = None
            for index in range(1, current.Folders.Count + 1):
                folder = current.Folders.Item(index)
                if str(folder.Name).strip().lower() == wanted:
                    match = folder
                    break
            if match is None:
                raise OutlookError(
                    f"Folder '{part}' not found under '{current.Name}'. "
                    "Check outlook.source_folders in config/settings.yaml."
                )
            current = match
        return current

    def _drafts(self):
        return self._namespace.GetDefaultFolder(OL_FOLDER_DRAFTS)

    def _ensure_folder(self, path: str, parent: str = "inbox"):
        """Find a folder path, creating missing levels.

        ``parent`` selects the root: "drafts" for folders that hold unsent mail,
        "inbox" for everything else. Unsent items belong under Drafts so Outlook
        keeps opening them in compose mode with a working Send button.
        """
        current = self._drafts() if parent == "drafts" else self._inbox()
        for part in path.replace("\\", "/").strip("/").split("/"):
            wanted = part.strip()
            match = None
            for index in range(1, current.Folders.Count + 1):
                folder = current.Folders.Item(index)
                if str(folder.Name).strip().lower() == wanted.lower():
                    match = folder
                    break
            if match is None:
                logger.info("Creating Outlook folder '%s' under '%s'", wanted, current.Name)
                match = current.Folders.Add(wanted)
            current = match
        return current

    def ensure_folder(self, path: str, parent: str = "inbox") -> str:
        """Create a folder path if missing. Returns the folder's display name."""
        return str(self._ensure_folder(path, parent).Name)

    def folder_exists(self, name: str) -> bool:
        try:
            self._find_folder(name)
            return True
        except OutlookError:
            return False

    def _walk(self, folder, path: str, recursive: bool):
        """Yield ``(folder, path)`` for a folder and, optionally, its descendants.

        Recursion is what lets a technician create a new client subfolder in
        Outlook and have it picked up without editing any configuration.
        """
        yield folder, path
        if not recursive:
            return
        for index in range(1, folder.Folders.Count + 1):
            child = folder.Folders.Item(index)
            yield from self._walk(child, f"{path}/{child.Name}", recursive)

    # ----------------------------------------------------------------- reading

    def messages(
        self, folder_name: str, include_subfolders: bool = True
    ) -> list[MailSummary]:
        """Snapshot every message in the folder tree as plain data.

        **Not** filtered by the read flag, deliberately. A technician who opens an
        alert to look at it marks it read, and filtering on unread would make that
        alarm invisible to the program — lost, silently. The ledger decides what
        has been handled; the read flag is only a cue for humans.

        The folder stays small because processed mail is moved to the processed
        folder, so re-reading everything costs nothing.

        The identifiers are collected up front on purpose: an Outlook collection
        updates live, so mutating ``UnRead`` while iterating shifts the indices and
        silently skips messages.
        """
        root = self._find_folder(folder_name)
        summaries: list[MailSummary] = []

        for folder, path in self._walk(root, folder_name, include_subfolders):
            items = folder.Items
            items.Sort("[ReceivedTime]", False)

            for index in range(1, items.Count + 1):
                item = items.Item(index)
                try:
                    if int(item.Class) != OL_CLASS_MAIL:
                        continue
                    entry_id = str(item.EntryID)
                    summaries.append(
                        MailSummary(
                            entry_id=entry_id,
                            message_key=_safe_property(item, PROP_INTERNET_MESSAGE_ID)
                            or entry_id,
                            subject=str(item.Subject or ""),
                            received=str(item.ReceivedTime),
                            sender=str(getattr(item, "SenderEmailAddress", "") or ""),
                            folder=path,
                        )
                    )
                except Exception as exc:  # noqa: BLE001
                    logger.warning(
                        "Skipping unreadable item %s in '%s': %s", index, path, exc
                    )

        return summaries

    def _item(self, entry_id: str):
        try:
            return self._namespace.GetItemFromID(entry_id)
        except Exception as exc:  # noqa: BLE001
            raise OutlookError(f"Message {entry_id} is no longer available") from exc

    def save_pdf_attachments(
        self, entry_id: str, dest_dir: Path, max_mb: int = 20
    ) -> list[Path]:
        """Save real PDF attachments, skipping inline signature images."""
        item = self._item(entry_id)
        dest_dir.mkdir(parents=True, exist_ok=True)
        saved: list[Path] = []

        attachments = item.Attachments
        for index in range(1, attachments.Count + 1):
            attachment = attachments.Item(index)
            name = str(attachment.FileName or "")

            # Extension is the only filter, deliberately. There was a Content-ID
            # check here to skip inline signature logos, and it silently dropped a
            # real report: Gmail sets a Content-ID on every attachment, PDFs
            # included, so the alarm was ignored as "no PDF attached". A
            # Content-ID means the part has an identifier, not that it is
            # decoration — and a signature logo is never a PDF anyway.
            if not name.lower().endswith(".pdf"):
                logger.debug("Ignoring non-PDF attachment %s", name)
                continue

            size_mb = int(getattr(attachment, "Size", 0) or 0) / (1024 * 1024)
            if size_mb > max_mb:
                logger.warning("Ignoring %s: %.1f MB exceeds the %s MB limit", name, size_mb, max_mb)
                continue

            destination = dest_dir / sanitize_filename(name, fallback=f"adjunto_{index}.pdf")
            attachment.SaveAsFile(str(destination))
            saved.append(destination)

        return saved

    # --------------------------------------------------------------- signature

    def _default_signature(self) -> str:
        """The account's default "new message" signature, as an HTML document.

        Classic Outlook only injects a signature when an Inspector is created
        for the item, which is why a throwaway ``MailItem`` is built and its
        ``GetInspector`` merely touched: that access is what makes Outlook drop
        the signature into ``HTMLBody``, where it can be read back. The item is
        never saved nor displayed, so nothing reaches the mailbox and it dies
        with this scope.

        Reading the signature from Outlook rather than hardcoding one in the
        template is deliberate: it keeps the logo exactly as the account has it.

        The HTML alone is not the signature. A logo is referenced as ``cid:...``
        and travels as a hidden attachment of the item it was inserted into, so
        copying only the HTML into another draft leaves a broken image, which in
        a signature that is just a picture means no signature at all. The
        referenced images are therefore saved here too, and ``_attach_signature_images``
        puts them on every draft that carries the signature.

        An account with no signature configured is an ordinary situation, not a
        failure, and neither is an Outlook build that refuses the Inspector. Any
        such case answers with an empty string and the draft is written exactly
        as it was before this existed.
        """
        if self._signature_cache is not None:
            return self._signature_cache

        signature = ""
        probe = None
        try:
            probe = self._app.CreateItem(OL_MAIL_ITEM)
            probe.GetInspector  # noqa: B018 - the access itself is the effect
            signature = str(probe.HTMLBody or "")
        except Exception as exc:  # noqa: BLE001 - COM raises bare com_error
            logger.info("Could not read the default Outlook signature: %s", exc)

        images: list[tuple[Path, str]] = []
        if signature.strip() and probe is not None:
            # Apart from the read above, so a picture that cannot be collected is
            # reported as that and never as a signature that could not be read.
            try:
                images = self._save_signature_images(probe, signature)
            except Exception as exc:  # noqa: BLE001
                logger.warning("Could not collect the signature's pictures: %s", exc)
        else:
            logger.debug("No default Outlook signature to add to the draft")

        self._signature_cache = signature
        self._signature_images = tuple(images)
        return signature

    def _save_signature_images(self, probe, signature: str) -> list[tuple[Path, str]]:
        """Save the pictures the signature references, with their Content-ID.

        The Content-IDs come from the HTML itself, since that is what a draft has
        to be able to resolve. The bytes come from Outlook's own attachments when
        it can hand them over, and from the picture embedded in the signature file
        when it cannot: some signatures keep their logo as a data URI, and Outlook
        refuses to save those ("the data source is missing").

        The folder is a fixed one under the system temp directory, so a run
        overwrites the files of the previous one instead of piling up new ones.
        """
        cids = list(dict.fromkeys(_CID_REFERENCE.findall(signature)))
        if not cids:
            return []

        folder = Path(tempfile.gettempdir()) / "clariot-signature"
        folder.mkdir(parents=True, exist_ok=True)

        saved: dict[str, Path] = {}
        try:
            count = int(probe.Attachments.Count)
        except Exception:  # noqa: BLE001
            count = 0
        for index in range(1, count + 1):
            try:
                attachment = probe.Attachments.Item(index)
                cid = str(attachment.PropertyAccessor.GetProperty(PROP_ATTACH_CONTENT_ID) or "")
                if cid not in cids:
                    continue
                path = folder / sanitize_filename(
                    str(attachment.FileName), fallback=f"firma-{index}"
                )
                attachment.SaveAsFile(str(path))
            except Exception as exc:  # noqa: BLE001 - not inline, or unsaveable
                logger.debug("Outlook would not hand over a signature picture: %s", exc)
                continue
            saved[cid] = path

        missing = [cid for cid in cids if cid not in saved]
        if missing:
            signatures = self._signatures_dir or (
                Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
            )
            embedded = _embedded_signature_images(signatures, _default_signature_name())
            if not embedded:
                logger.info(
                    "The signature has %s picture(s) that could not be recovered",
                    len(missing),
                )
            # Matched by order of appearance: the k-th picture the HTML references
            # is the k-th one stored in the file.
            for number, (cid, (data, kind)) in enumerate(zip(missing, embedded), start=1):
                path = folder / f"firma-{number}.{kind}"
                path.write_bytes(data)
                saved[cid] = path

        return [(path, cid) for cid, path in saved.items()]

    def _attach_signature_images(self, mail) -> None:
        """Give a draft the hidden images its signature points at.

        Each is attached under the Content-ID the signature HTML already uses, so
        the ``cid:`` reference resolves. A failure only costs the picture, never
        the draft.
        """
        for path, cid in self._signature_images:
            try:
                attachment = mail.Attachments.Add(str(path))
                attachment.PropertyAccessor.SetProperty(PROP_ATTACH_CONTENT_ID, cid)
                attachment.PropertyAccessor.SetProperty(PROP_ATTACH_HIDDEN, True)
            except Exception as exc:  # noqa: BLE001 - COM raises bare com_error
                logger.warning("Could not attach the signature image %s: %s", path.name, exc)

    def _signature(self) -> str:
        """The signature to merge, or "" when the setting switches it off."""
        if not self._use_signature:
            return ""
        return self._default_signature()

    # ----------------------------------------------------------------- writing

    def create_draft(
        self,
        *,
        subject: str,
        html_body: str,
        to: Sequence[str],
        cc: Sequence[str],
        attachments: Sequence[Path],
        sender_account: str = "",
        target_folder: str = "",
        target_folder_parent: str = "drafts",
    ) -> str:
        """Create a draft and save it. Never sends. Returns the draft EntryID."""
        mail = self._app.CreateItem(OL_MAIL_ITEM)
        mail.Subject = subject
        if to:
            mail.To = "; ".join(to)
        if cc:
            mail.CC = "; ".join(cc)

        # Before the signature is read, on purpose: which signature Outlook
        # offers depends on the account the message is being sent from.
        if sender_account:
            account = self._account_by_smtp(sender_account)
            if account is not None:
                mail.SendUsingAccount = account
            else:
                logger.warning("Account '%s' not found; using the default one", sender_account)

        # The signature's pictures go on first and the body is written after them:
        # Outlook only links a cid: reference to an attachment that already exists
        # when the HTML arrives, otherwise the picture lands as a plain file.
        signature = self._signature()
        self._attach_signature_images(mail)
        mail.HTMLBody = _merge_signature(html_body, signature)

        for path in attachments:
            if path.exists():
                # Outlook rejects relative paths and forward slashes.
                mail.Attachments.Add(str(path.resolve()))
            else:
                logger.error("Attachment missing at draft time: %s", path)

        mail.Save()
        if target_folder:
            # Move returns a new item; its EntryID differs from the saved one.
            mail = mail.Move(self._ensure_folder(target_folder, target_folder_parent))
        return str(mail.EntryID)

    # Written by us into every draft. Its absence means a human edited the body.
    UNTOUCHED_MARKER = "[COMPLETAR]"

    DRAFT_MISSING = "missing"
    DRAFT_SENT = "sent"
    DRAFT_TOUCHED = "touched"
    DRAFT_CLEAN = "clean"

    def draft_state(
        self,
        entry_id: str,
        *,
        expected_cc: Sequence[str] = (),
        expected_to: Sequence[str] = (),
    ) -> str:
        """Whether a draft can safely be regenerated.

        Regenerating destroys whatever the technician typed, so a draft he has
        started working on is left alone and the new alarm gets its own draft.

        ``expected_cc`` is the copy list the program itself would have written for
        this draft. Only a *longer* CC than that means a human added a recipient;
        the copies we put there ourselves are not evidence of anything.

        ``expected_to`` is the same idea for the recipients: the addresses the
        client directory pre-loads. With nothing pre-loaded, any recipient at all
        was typed by a person.
        """
        try:
            item = self._namespace.GetItemFromID(entry_id)
        except Exception:  # noqa: BLE001 - deleted, or the id no longer resolves
            return self.DRAFT_MISSING

        try:
            if bool(getattr(item, "Sent", False)):
                return self.DRAFT_SENT
            # Counted, never matched by address. After Save(), Outlook resolves a
            # recipient against the address book and reads the entry back as a
            # display name ("Orellana, Pablo") instead of the SMTP address we wrote,
            # so comparing strings would see our own recipients as a human edit and
            # every draft would stop being reusable — which is the bug this check
            # replaces. The number of entries survives that rewriting.
            if _recipient_count(str(item.To or "")) > len(expected_to):
                return self.DRAFT_TOUCHED
            if _recipient_count(str(item.CC or "")) > len(expected_cc):
                return self.DRAFT_TOUCHED
            if self.UNTOUCHED_MARKER not in str(item.HTMLBody or ""):
                return self.DRAFT_TOUCHED
        except Exception:  # noqa: BLE001
            return self.DRAFT_MISSING
        return self.DRAFT_CLEAN

    def update_draft(
        self,
        entry_id: str,
        *,
        subject: str,
        html_body: str,
        cc: Sequence[str] = (),
        attachments: Sequence[Path] = (),
        target_folder: str = "",
        target_folder_parent: str = "drafts",
    ) -> str:
        """Rewrite an existing draft in place. Returns its (possibly new) EntryID.

        Attachments are replaced rather than appended, so re-running the same
        alarm set does not pile up duplicate PDFs.
        """
        item = self._item(entry_id)
        item.Subject = subject
        if cc:
            item.CC = "; ".join(cc)

        for index in range(item.Attachments.Count, 0, -1):
            item.Attachments.Item(index).Delete()
        # The loop above also removed the signature's hidden images, so they go
        # back on before the body is written (see create_draft). The body is
        # merged again, not carried over: it is rewritten wholesale, so a draft
        # growing a second alarm would otherwise lose its signature.
        signature = self._signature()
        self._attach_signature_images(item)
        item.HTMLBody = _merge_signature(html_body, signature)
        for path in attachments:
            if path.exists():
                item.Attachments.Add(str(path.resolve()))
            else:
                logger.error("Attachment missing while updating a draft: %s", path)

        item.Save()
        if target_folder:
            current = str(item.Parent.Name).strip().lower()
            if current != target_folder.split("/")[-1].strip().lower():
                # Severity climbed, so the draft moves folder. Move returns a new
                # object whose EntryID differs.
                item = item.Move(self._ensure_folder(target_folder, target_folder_parent))
        return str(item.EntryID)

    def _account_by_smtp(self, smtp: str):
        accounts = self._namespace.Accounts
        for index in range(1, accounts.Count + 1):
            account = accounts.Item(index)
            if str(getattr(account, "SmtpAddress", "") or "").lower() == smtp.lower():
                return account
        return None

    def inject_test_message(
        self,
        folder_name: str,
        subject: str,
        attachments: Sequence[Path],
        body: str = "",
        message_id: str = "",
    ) -> str:
        """Drop an unread message straight into a folder, for testing.

        Preferred over actually emailing yourself: nothing leaves the mailbox, no
        Outlook rule has to exist yet, and it is instant. Delete the item to undo.
        """
        folder = self._ensure_folder(folder_name)
        item = self._app.CreateItem(OL_MAIL_ITEM)
        item.Subject = subject
        item.Body = body or "Mensaje de prueba generado por clariot --simulate."
        for path in attachments:
            if not path.exists():
                raise OutlookError(f"Attachment not found: {path}")
            # Outlook rejects relative paths and forward slashes.
            item.Attachments.Add(str(path.resolve()))

        # Saving an unsent item always files it under Drafts, whatever folder it
        # was created in, so it has to be moved afterwards. Move returns a new
        # object; the read flag is set on that one.
        item.Save()
        moved = item.Move(folder)

        # Real mail carries an Internet Message-ID that survives folder moves,
        # and the ledger keys on it. Without one, a locally created item falls
        # back to its EntryID, which changes on every move — so the test fixture
        # would not exercise the duplicate guard faithfully.
        if message_id:
            try:
                moved.PropertyAccessor.SetProperty(PROP_INTERNET_MESSAGE_ID, message_id)
            except Exception as exc:  # noqa: BLE001
                logger.warning("Could not set a synthetic Message-ID: %s", exc)

        moved.UnRead = True
        moved.Save()
        return str(moved.EntryID)

    def finish_message(
        self, entry_id: str, *, mark_read: bool = True, move_to: str = ""
    ) -> None:
        """Mark the original as read and optionally file it away."""
        item = self._item(entry_id)
        if mark_read:
            item.UnRead = False
            item.Save()
        if move_to:
            item.Move(self._ensure_folder(move_to))
