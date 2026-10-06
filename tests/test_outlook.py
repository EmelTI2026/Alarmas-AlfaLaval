"""``draft_state`` decides whether a draft may be rewritten.

No real Outlook is involved. ``OutlookClient.__init__`` dispatches COM, which
only exists on a Windows machine running classic Outlook, so the instance is
built without it and given a namespace that hands back one fake mail item —
everything ``draft_state`` actually reads.
"""

from types import SimpleNamespace

from clariot.adapters.outlook import OutlookClient

BODY = f"<html><p>Estimado {OutlookClient.UNTOUCHED_MARKER}</p></html>"

# The four copies config/clients.yaml puts on every draft.
ALWAYS_CC = (
    "porellana@emeltec.cl",
    "cfarias@emeltec.cl",
    "mlizana@emeltec.cl",
    "hlagos@emeltec.cl",
)

# How Outlook reads those same four back after Save(): resolved against the
# address book, so display names instead of addresses.
RESOLVED_CC = "Orellana, Pablo; Farias, Carlos; Lizana, Marco; Lagos, Hector"


def client(**item) -> OutlookClient:
    """An OutlookClient whose namespace returns one fake draft."""
    fields = {"Sent": False, "To": "", "CC": "", "HTMLBody": BODY}
    fields.update(item)
    draft = SimpleNamespace(**fields)

    adapter = object.__new__(OutlookClient)
    adapter._namespace = SimpleNamespace(GetItemFromID=lambda entry_id: draft)
    return adapter


def test_our_own_copies_are_not_evidence_that_a_human_edited_the_draft():
    """The regression: always_cc made every draft we created look edited.

    ``create_draft`` writes the copy list itself, so a non-empty CC used to mean
    the draft was untouchable from the very run that produced it. The second
    alarm on a machine then opened a separate draft instead of growing the open
    one, taking the severity escalation down with it.
    """
    adapter = client(CC=RESOLVED_CC)

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC) == adapter.DRAFT_CLEAN


def test_an_extra_copy_means_a_human_edited_the_draft():
    adapter = client(CC=f"{RESOLVED_CC}; jefatura@cliente.example")

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC) == adapter.DRAFT_TOUCHED


def test_a_display_name_with_a_comma_counts_as_one_recipient():
    """Outlook separates with ';'. Splitting on every ',' would re-break reuse."""
    adapter = client(CC="Orellana, Pablo")

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC[:1]) == adapter.DRAFT_CLEAN


def test_two_addresses_typed_on_one_line_count_as_two():
    """A comma only splits a piece that plainly holds two addresses."""
    adapter = client(CC="uno@cliente.example, dos@cliente.example")

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC[:1]) == adapter.DRAFT_TOUCHED


def test_a_recipient_in_the_to_field_still_means_a_human_edited_it():
    """Drafts are created without a recipient; the technician adds it by hand."""
    adapter = client(To="operaciones@cliente.example", CC=RESOLVED_CC)

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC) == adapter.DRAFT_TOUCHED


def test_a_body_without_the_placeholder_means_a_human_edited_it():
    adapter = client(CC=RESOLVED_CC, HTMLBody="<html><p>Estimado Pablo</p></html>")

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC) == adapter.DRAFT_TOUCHED


def test_a_sent_draft_is_never_rewritten():
    adapter = client(Sent=True, CC=RESOLVED_CC)

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC) == adapter.DRAFT_SENT


def test_expected_cc_defaults_to_nothing_so_old_callers_keep_working():
    """A caller that passes no copy list still gets the previous meaning."""
    assert client(CC="").draft_state("ENTRY-1") == OutlookClient.DRAFT_CLEAN
    assert client(CC="alguien@cliente.example").draft_state("ENTRY-1") == (
        OutlookClient.DRAFT_TOUCHED
    )


def test_a_draft_that_no_longer_resolves_is_reported_missing():
    adapter = object.__new__(OutlookClient)

    def gone(entry_id):
        raise RuntimeError("the item was deleted")

    adapter._namespace = SimpleNamespace(GetItemFromID=gone)

    assert adapter.draft_state("ENTRY-1", expected_cc=ALWAYS_CC) == adapter.DRAFT_MISSING
