"""Testes das funções puras do scripts/build_union.py."""

import build_union as bu
import pytest


# ── parse_amount ────────────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "valor, esperado",
    [
        ("€1,234.56", 1234.56),
        ("R$379.71", 379.71),
        ("250477.99", 250477.99),
        ("-13.84", -13.84),
        ("", 0.0),
        ("nan", 0.0),
        (None, 0.0),
        ("N/A", 0.0),
    ],
)
def test_parse_amount(valor, esperado):
    assert bu.parse_amount(valor) == pytest.approx(esperado)


# ── normalize_login ─────────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "username, esperado",
    [
        ("affiliabetbrasil@gmail.com", "affiliabetbrasil"),
        ("info@afiliagambling.com", "infoafiliagambling"),
        ("infoafiliawin@gmail.com", "infoafiliawin"),
        ("TipsterpageBR", "TipsterpageBR"),  # não-e-mail fica igual
        ("Affiliabet_Panama", "Affiliabet_Panama"),
    ],
)
def test_normalize_login(username, esperado):
    assert bu.normalize_login(username) == esperado


# ── total_facturacion ───────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "resta, cpa, rs, esperado",
    [
        # Resta = No: ignora RS negativo
        ("No", 6440.0, -1601.38, 6440.0),
        ("No", 0.0, 379.71, 379.71),
        ("No", 0.0, -2.96, 0.0),
        ("No", 35.0, 0.0, 35.0),
        # Resta = Sí: soma; se soma < 0 -> 0; se RS<0 e CPA=0 -> CPA
        ("Sí", 8710.0, -126.34, 8583.66),
        ("Sí", 10.0, -50.0, 0.0),       # soma negativa -> 0
        ("Sí", 0.0, -5.0, 0.0),         # RS<0 e CPA=0 -> CPA+0 = 0
        ("Sí", 100.0, 50.0, 150.0),
        # Outros valores de Resta: CPA + RS direto
        ("", 10.0, -3.0, 7.0),
        ("qualquer", 10.0, 20.0, 30.0),
        # Não-numérico -> 0 (try/otherwise)
        ("No", "x", 5.0, 0.0),
        ("Sí", None, None, 0.0),
    ],
)
def test_total_facturacion(resta, cpa, rs, esperado):
    assert bu.total_facturacion(resta, cpa, rs) == pytest.approx(esperado)


# ── _to_eur ─────────────────────────────────────────────────────────────────
def test_to_eur_converte():
    rates = {("2026-05", "BRL"): 5.8872, ("2026-05", "EUR"): 1.0}
    assert bu._to_eur(379.71, "2026-05", "BRL", rates) == pytest.approx(64.50, abs=0.01)
    assert bu._to_eur(100.0, "2026-05", "EUR", rates) == 100.0


def test_to_eur_sem_taxa_retorna_none():
    assert bu._to_eur(10.0, "2026-05", "XXX", {}) is None


def test_to_eur_taxa_zero_retorna_none():
    assert bu._to_eur(10.0, "2026-05", "BRL", {("2026-05", "BRL"): 0}) is None
