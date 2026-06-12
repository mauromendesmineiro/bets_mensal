"""Testes das funções puras do scripts/common.py."""

from datetime import date

import common
import pandas as pd
import pytest


# ── sanitize / env_key_for ──────────────────────────────────────────────────
@pytest.mark.parametrize(
    "val, esperado",
    [
        ("BetsAmigo", "BETSAMIGO"),
        ("info@afiliagambling.com", "INFO_AFILIAGAMBLING_COM"),
        ("Tipster page-BR", "TIPSTER_PAGE_BR"),
    ],
)
def test_sanitize(val, esperado):
    assert common.sanitize(val) == esperado


def test_env_key_for():
    assert (
        common.env_key_for("CELLXPERT", "BetsAmigo", "AffiliabetLATAM")
        == "PASS_CELLXPERT_BETSAMIGO_AFFILIABETLATAM"
    )


# ── previous_month_label ────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "hoje, esperado",
    [
        (date(2026, 6, 11), "2026-05"),
        (date(2026, 1, 5), "2025-12"),
        (date(2026, 3, 1), "2026-02"),
    ],
)
def test_previous_month_label(hoje, esperado):
    assert common.previous_month_label(hoje) == esperado


# ── safe_name ───────────────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "nome, esperado",
    [
        ("Novibet / info", "Novibet___info"),
        ("a:b*c", "a_b_c"),
        ("", "report"),
        ("___", "report"),
    ],
)
def test_safe_name(nome, esperado):
    assert common.safe_name(nome) == esperado


# ── detect_currency ─────────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "valor, esperado",
    [
        ("€21.47", "EUR"),
        ("R$379.71", "BRL"),
        ("£10", "GBP"),
        ("$5", "USD"),
        ("MXN 100", "MXN"),
        ("1234.56", None),
        (None, None),
        ("", None),
    ],
)
def test_detect_currency(valor, esperado):
    assert common.detect_currency(valor) == esperado


def test_detect_currency_brl_antes_de_usd():
    # 'R$' deve ganhar de '$' (ordem dos prefixos importa).
    assert common.detect_currency("R$5") == "BRL"


# ── canonical_operador / norm_key ───────────────────────────────────────────
@pytest.mark.parametrize(
    "operador, esperado",
    [
        ("Sportium", "SportiumBet"),
        ("sportium", "SportiumBet"),
        ("Betano", "Betano"),  # sem alias fica igual
        (None, ""),
    ],
)
def test_canonical_operador(operador, esperado):
    assert common.canonical_operador(operador) == esperado


@pytest.mark.parametrize(
    "val, esperado",
    [("  Afiliagambling ", "afiliagambling"), ("BETFAIR", "betfair"), (None, "")],
)
def test_norm_key(val, esperado):
    assert common.norm_key(val) == esperado


def test_retry_until_sucesso_na_segunda():
    estado = {"n": 0}

    def fn():
        estado["n"] += 1
        return None if estado["n"] < 2 else "ok"

    r = common.retry_until(fn, ok=lambda x: x is not None, attempts=3, backoff_s=0)
    assert r == "ok"
    assert estado["n"] == 2


def test_retry_until_esgota_tentativas():
    estado = {"n": 0}

    def fn():
        estado["n"] += 1
        return None

    r = common.retry_until(fn, ok=lambda x: x is not None, attempts=3, backoff_s=0)
    assert r is None
    assert estado["n"] == 3


def test_previous_month_label_target_month_override(monkeypatch):
    monkeypatch.setenv("TARGET_MONTH", "2026-03")
    assert common.previous_month_label() == "2026-03"
    # today explícito ignora o override (mantém determinismo)
    assert common.previous_month_label(date(2026, 6, 11)) == "2026-05"


# ── upsert_csv ──────────────────────────────────────────────────────────────
def test_upsert_csv_substitui_periodo(tmp_path):
    path = tmp_path / "u.csv"
    base = pd.DataFrame(
        {"operador": ["A", "A"], "month": ["2026-04", "2026-05"], "v": ["10", "20"]}
    )
    base.to_csv(path, index=False, encoding="utf-8-sig")

    novo = pd.DataFrame({"operador": ["A"], "month": ["2026-05"], "v": ["99"]})
    out = common.upsert_csv(path, novo, ["operador", "month"])

    # 2026-04 preservado, 2026-05 substituído (sem duplicar)
    assert len(out) == 2
    assert set(out["month"]) == {"2026-04", "2026-05"}
    assert out[out["month"] == "2026-05"]["v"].iloc[0] == "99"
    assert out[out["month"] == "2026-04"]["v"].iloc[0] == "10"


def test_upsert_csv_replace_keys_remove_orfaos(tmp_path):
    path = tmp_path / "u.csv"
    base = pd.DataFrame(
        {"operador": ["A", "B"], "month": ["2026-05", "2026-05"], "v": ["1", "2"]}
    )
    base.to_csv(path, index=False, encoding="utf-8-sig")

    # Lote novo só tem A; replace_keys inclui B (processado mas removido) -> B sai.
    novo = pd.DataFrame({"operador": ["A"], "month": ["2026-05"], "v": ["10"]})
    replace = pd.DataFrame({"operador": ["A", "B"], "month": ["2026-05", "2026-05"]})
    out = common.upsert_csv(path, novo, ["operador", "month"], replace_keys=replace)

    assert list(out["operador"]) == ["A"]
    assert out["v"].iloc[0] == "10"


def test_upsert_csv_cria_quando_nao_existe(tmp_path):
    path = tmp_path / "novo.csv"
    novo = pd.DataFrame({"operador": ["A"], "month": ["2026-05"], "v": ["1"]})
    out = common.upsert_csv(path, novo, ["operador", "month"])
    assert len(out) == 1
