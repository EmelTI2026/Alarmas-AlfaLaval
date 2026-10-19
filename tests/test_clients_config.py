"""The shipped config/clients.yaml routes each client to the right people.

These tests read the real file, not a fixture: they exist so that an edit to a
recipient list cannot silently send a client's report to the wrong place.
"""

from pathlib import Path

import pytest

from clariot.config import ClientDirectory

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"

EMELTEC_COPIES = {
    "porellana@emeltec.cl",
    "cfarias@emeltec.cl",
    "hlagos@emeltec.cl",
    "mlizana@emeltec.cl",
}


@pytest.fixture(scope="module")
def directory() -> ClientDirectory:
    return ClientDirectory.load(CONFIG_DIR)


# What the PDF puts in the company field, as seen in real reports.
RECIPIENTS = {
    "Prolesur": {"guillermo.lorcat@soprole.cl"},
    "Vital Jugos": {"rletelier@koandina.com", "mvalenzuelal@koandina.com"},
    "Soprole": {"hernan.vergara@soprole.cl", "luis.olea@soprole.cl"},
    "Surlat": {
        "fmatamala@quillayessurlat.cl",
        "hcortes@quillayessurlat.cl",
        "jtorres@quillayessurlat.cl",
        "edonoso@quillayessurlat.cl",
    },
}


@pytest.mark.parametrize("company", RECIPIENTS)
def test_each_client_gets_exactly_its_own_recipients(directory, company):
    route = directory.resolve(company)

    assert route is not None, f"{company} matches no client"
    assert set(route.to) == RECIPIENTS[company]


def test_the_plant_field_resolves_like_the_company(directory):
    route = directory.resolve("Soprole - San Bernardo Planta 4")

    assert route is not None
    assert set(route.to) == RECIPIENTS["Soprole"]


def test_prolesur_is_not_mistaken_for_soprole(directory):
    """Both are routed to soprole.cl addresses, but they are different lists."""
    assert set(directory.resolve("Prolesur").to) != set(directory.resolve("Soprole").to)


def test_the_four_emeltec_copies_go_on_every_draft(directory):
    assert set(directory.always_cc) == EMELTEC_COPIES


def test_the_extra_copy_is_only_on_soprole(directory):
    assert directory.resolve("Soprole").cc == ("cvarela@emeltec.cl",)
    for company in ("Prolesur", "Vital Jugos", "Surlat"):
        assert "cvarela@emeltec.cl" not in directory.resolve(company).cc
    assert "cvarela@emeltec.cl" not in directory.always_cc


def test_an_unknown_company_still_gets_an_empty_recipient(directory):
    assert directory.resolve("Empresa Nueva") is None
