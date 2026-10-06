"""Outlook desktop integration through COM (pywin32).

Deployment constraints worth knowing before touching this file:

* Only *classic* Outlook for Windows exposes a COM automation surface. The new
  Outlook client does not, and no amount of patching here will change that.
* COM requires an interactive desktop session. A Scheduled Task configured with
  "Run whether user is logged on or not" fails under session 0 isolation. Use an
  "on logon" trigger with repetition instead, on a machine that stays signed in.
"""

from __future__ import annotations

import logging
import re
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

    def __init__(self, *, use_signature: bool = True) -> None:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:  # pragma: no cover - Windows-only dependency
            raise OutlookError("pywin32 is required to talk to Outlook") from exc

        self._use_signature = use_signature
        # Read at most once per run, and only if a draft is actually written.
        self._signature_cache: str | None = None

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
        template is deliberate: it keeps the logo and the inline images exactly
        as the account has them.

        An account with no signature configured is an ordinary situation, not a
        failure, and neither is an Outlook build that refuses the Inspector. Any
        such case answers with an empty string and the draft is written exactly
        as it was before this existed.
        """
        if self._signature_cache is not None:
            return self._signature_cache

        signature = ""
        try:
            probe = self._app.CreateItem(OL_MAIL_ITEM)
            probe.GetInspector  # noqa: B018 - the access itself is the effect
            signature = str(probe.HTMLBody or "")
        except Exception as exc:  # noqa: BLE001 - COM raises bare com_error
            logger.info("Could not read the default Outlook signature: %s", exc)

        if not signature.strip():
            logger.debug("No default Outlook signature to add to the draft")
        self._signature_cache = signature
        return signature

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

        mail.HTMLBody = _merge_signature(html_body, self._signature())

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

    def draft_state(self, entry_id: str) -> str:
        """Whether a draft can safely be regenerated.

        Regenerating destroys whatever the technician typed, so a draft he has
        started working on is left alone and the new alarm gets its own draft.
        """
        try:
            item = self._namespace.GetItemFromID(entry_id)
        except Exception:  # noqa: BLE001 - deleted, or the id no longer resolves
            return self.DRAFT_MISSING

        try:
            if bool(getattr(item, "Sent", False)):
                return self.DRAFT_SENT
            if str(item.To or "").strip() or str(item.CC or "").strip():
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
        # Merged again, not carried over: the body is rewritten wholesale, so a
        # draft growing a second alarm would otherwise lose its signature.
        item.HTMLBody = _merge_signature(html_body, self._signature())
        if cc:
            item.CC = "; ".join(cc)

        for index in range(item.Attachments.Count, 0, -1):
            item.Attachments.Item(index).Delete()
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
