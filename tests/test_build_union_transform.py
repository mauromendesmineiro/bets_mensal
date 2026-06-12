"""Testes orientados a dados das transformações do scripts/build_union.py.

Usam DataFrames sintéticos (monkeypatch de _read) e um ref.xlsx temporário —
sem rede nem browser.
"""

import build_union as bu
import pandas as pd
import pytest


# ── from_netrefer: filtra Month == month ────────────────────────────────────
def test_from_netrefer_filtra_mes_do_periodo(monkeypatch):
    df = pd.DataFrame(
        {
            "plataforma": ["Netrefer", "Netrefer"],
            "operador": ["BetMGM", "BetMGM"],
            "empresa": ["Brasil", "Brasil"],
            "username": ["TipsterpageBR", "TipsterpageBR"],
            "month": ["2026-05", "2026-05"],
            "currency": ["BRL", "BRL"],
            "Month": ["2026-05", "2026-04"],          # só a 1ª casa o período
            "Revenue Share Reward": ["R$379.71", "R$10.00"],
            "CPA Reward": ["R$0.00", "R$0.00"],
        }
    )
    monkeypatch.setattr(bu, "_read", lambda name: df)
    out = bu.from_netrefer()
    assert len(out) == 1
    assert out["rs_operador"].iloc[0] == pytest.approx(379.71)
    assert out["cpa_operador"].iloc[0] == 0.0


# ── from_cellxpert: agrega CPA% e Revshare% ─────────────────────────────────
def test_from_cellxpert_agrega_por_tipo(monkeypatch):
    df = pd.DataFrame(
        {
            "plataforma": ["Cellxpert"] * 3,
            "operador": ["BetsAmigo"] * 3,
            "empresa": ["Affiliabet"] * 3,
            "username": ["AffiliabetLATAM"] * 3,
            "month": ["2026-05"] * 3,
            "currency": ["EUR"] * 3,
            "commission_type": ["CPA (Mexico)", "Revshare Ongoing NGR", "Fee X"],
            "amount": ["40.0", "2.79", "5.0"],
        }
    )
    monkeypatch.setattr(bu, "_read", lambda name: df)
    out = bu.from_cellxpert()
    assert len(out) == 1
    assert out["cpa_operador"].iloc[0] == pytest.approx(40.0)   # só CPA%
    assert out["rs_operador"].iloc[0] == pytest.approx(2.79)    # só Revshare%
    # 'Fee X' não entra em nenhum dos dois.


# ── from_income_access: remove duplicatas exactas ───────────────────────────
def test_from_income_access_dedup(monkeypatch):
    base = {
        "plataforma": "IncomeAccess",
        "operador": "Yosport",
        "empresa": "Afiliagambling",
        "account_username": "ReyAnalista",
        "month": "2026-05",
        "currency": "EUR",
        "pct_commission": "0.0",
        "cpa_commission": "7440.0",
    }
    df = pd.DataFrame([base, base])  # 2 linhas idênticas (rowid 1 e 2)
    monkeypatch.setattr(bu, "_read", lambda name: df)
    out = bu.from_income_access()
    assert len(out) == 1
    assert out["cpa_operador"].iloc[0] == pytest.approx(7440.0)


# ── enrich_ref: join normal + excepção Novibet/Afiliagambling via vendor_name ─
def test_enrich_ref_excecao_novibet(monkeypatch, tmp_path):
    ref = pd.DataFrame(
        {
            "operador": ["Betano", "Novibet Mexico"],
            "username": ["AG2022", "info@afiliagambling.com"],
            "empresa": ["Afiliagambling", "Afiliagambling"],
            "ref": ["BETANO ECUADOR", "NOVIBET MEXICO - info"],
            "resta": ["No", "Sí"],
        }
    )
    ref_path = tmp_path / "ref.xlsx"
    ref.to_excel(ref_path, index=False)
    monkeypatch.setattr(bu, "REF_XLSX", ref_path)

    union = pd.DataFrame(
        {
            "operador": ["Betano", "Novibet"],
            "username": ["AG2022", "info@afiliagambling.com"],
            "empresa": ["Afiliagambling", "Afiliagambling"],
            "vendor_name": ["", "Novibet Mexico"],  # excepção usa o vendor_name
        }
    )
    out = bu.enrich_ref(union)
    # Betano casa pela chave normal
    assert out.loc[out["operador"] == "Betano", "ref"].iloc[0] == "BETANO ECUADOR"
    # Novibet/Afiliagambling casa pelo vendor_name
    novibet = out[out["operador"] == "Novibet"].iloc[0]
    assert novibet["ref"] == "NOVIBET MEXICO - info"
    assert novibet["resta"] == "Sí"
