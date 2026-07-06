"""
scripts/build_union.py

Junta os relatórios por plataforma da pasta report/ num único ficheiro
report/union_data.csv com um esquema comum.

Colunas comuns (presentes em todas as plataformas):
    plataforma, operador, empresa, username, month, currency

Colunas derivadas:
    vendor_name   — só RavenTrack (uma linha por vendor); vazio nas restantes
    rs_operador   — comissão de revenue share do operador (numérico)
    cpa_operador  — comissão de CPA do operador (numérico)

Mapeamento por plataforma:
- netrefer.csv : filtra as linhas em que a coluna 'Month' == 'month' (o mês do
                 período) e usa 'Revenue Share Reward' → rs_operador e
                 'CPA Reward' → cpa_operador.
- raventrack.csv: uma linha por 'Vendor Name'; 'RevShare Commission' → rs_operador
                 e 'CPA Commission' → cpa_operador.
- cellxpert.csv: agrega por conta; soma 'amount' de commission_type que começa por
                 'CPA' → cpa_operador e por 'Revshare' → rs_operador.
- income_access.csv: 'pct_commission' → rs_operador e 'cpa_commission' →
                 cpa_operador (o username vem de 'account_username').

Os valores monetários (ex.: 'R$379.71', '€1,234.56', '250477.99') são
normalizados para número; a moeda fica na coluna 'currency'.

Uso:
    python scripts/build_union.py
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

try:  # funciona tanto em execução directa quanto importado como scripts.*
    import common
except ImportError:
    from scripts import common

ROOT = common.ROOT
REPORT_DIR = common.REPORT_DIR
RATES_CSV = ROOT / "data" / "currency_rates.csv"

log = common.get_logger("build_union")

COMMON_COLS = ["plataforma", "operador", "empresa", "username", "month", "currency"]
# vendor_name é usado apenas internamente (join com o ref para Novibet/Afiliagambling)
# e por isso não consta no output final.
OUTPUT_COLS = ["index"] + COMMON_COLS + [
    "rate",
    "rs_operador",
    "cpa_operador",
    "rs_eur",
    "cpa_eur",
    "ref",
    "resta",
]

REF_XLSX = ROOT / "config" / "ref.xlsx"

# Aliases de operador centralizados no common (ver common.OPERADOR_ALIASES).
OPERADOR_ALIASES = common.OPERADOR_ALIASES


def parse_amount(val) -> float:
    """Converte '€1,234.56' / 'R$379.71' / '250477.99' / '-13.84' em float."""
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s or s.lower() in {"nan", "n/a", "none"}:
        return 0.0
    s = s.replace(",", "")            # separador de milhares
    s = re.sub(r"[^0-9.\-]", "", s)   # remove símbolos de moeda e letras
    if s in {"", "-", ".", "-."}:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _read(name: str) -> pd.DataFrame | None:
    path = REPORT_DIR / name
    if not path.exists():
        log.warning(f"{name} não encontrado — ignorado")
        return None
    return pd.read_csv(path, dtype=str, encoding="utf-8-sig")


def _base_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Garante as colunas comuns + as derivadas, na ordem final."""
    out = pd.DataFrame()
    for c in COMMON_COLS:
        out[c] = df[c] if c in df.columns else ""
    out["vendor_name"] = ""
    out["rs_operador"] = 0.0
    out["cpa_operador"] = 0.0
    return out


def from_netrefer() -> pd.DataFrame:
    df = _read("netrefer.csv")
    if df is None or df.empty:
        return pd.DataFrame(columns=OUTPUT_COLS)
    # Mantém apenas a linha do mês do período (Month == month).
    df = df[df["Month"].astype(str).str.strip() == df["month"].astype(str).str.strip()]
    out = _base_frame(df)
    out["rs_operador"] = df["Revenue Share Reward"].map(parse_amount).values
    out["cpa_operador"] = df["CPA Reward"].map(parse_amount).values
    # 22Bet: ao contrário dos restantes operadores da Netrefer, o CPA correcto
    # vem de 'Reward Due' (não de 'CPA Reward') e o RS não entra na conta (fica 0).
    is_22bet = df["operador"].astype(str).str.strip() == "22Bet"
    if is_22bet.any():
        out.loc[is_22bet.values, "cpa_operador"] = (
            df.loc[is_22bet, "Reward Due"].map(parse_amount).values
        )
        out.loc[is_22bet.values, "rs_operador"] = 0.0
    return out.reset_index(drop=True)


def from_raventrack() -> pd.DataFrame:
    df = _read("raventrack.csv")
    if df is None or df.empty:
        return pd.DataFrame(columns=OUTPUT_COLS)
    out = _base_frame(df)
    # No RavenTrack a moeda real dos valores está na coluna 'Currency' (da tabela),
    # não na meta 'currency' (que reflecte a moeda da conta, ex.: USD/EUR).
    if "Currency" in df.columns:
        out["currency"] = df["Currency"].values
    out["vendor_name"] = df["Vendor Name"].fillna("").values
    out["rs_operador"] = df["RevShare Commission"].map(parse_amount).values
    out["cpa_operador"] = df["CPA Commission"].map(parse_amount).values
    return out.reset_index(drop=True)


def from_cellxpert() -> pd.DataFrame:
    df = _read("cellxpert.csv")
    if df is None or df.empty:
        return pd.DataFrame(columns=OUTPUT_COLS)
    df = df.copy()
    df["amount_num"] = df["amount"].map(parse_amount)
    ct = df["commission_type"].fillna("").astype(str).str.strip().str.lower()
    df["is_cpa"] = ct.str.startswith("cpa")
    df["is_rs"] = ct.str.startswith("revshare")
    df["cpa_part"] = df["amount_num"].where(df["is_cpa"], 0.0)
    df["rs_part"] = df["amount_num"].where(df["is_rs"], 0.0)

    grp = (
        df.groupby(COMMON_COLS, dropna=False)[["cpa_part", "rs_part"]]
        .sum()
        .reset_index()
    )
    out = _base_frame(grp)
    out["cpa_operador"] = grp["cpa_part"].values
    out["rs_operador"] = grp["rs_part"].values
    return out.reset_index(drop=True)


def from_income_access() -> pd.DataFrame:
    df = _read("income_access.csv")
    if df is None or df.empty:
        return pd.DataFrame(columns=OUTPUT_COLS)
    df = df.copy()
    # No income_access o username está em 'account_username'.
    if "username" not in df.columns and "account_username" in df.columns:
        df["username"] = df["account_username"]
    out = _base_frame(df)
    out["rs_operador"] = df["pct_commission"].map(parse_amount).values
    out["cpa_operador"] = df["cpa_commission"].map(parse_amount).values
    # O income_access exporta 2 linhas por conta: o registo real (rowid=1) e uma
    # linha-"total" que repete os mesmos valores. São idênticas nos campos do union,
    # por isso removemos as duplicatas exactas para não duplicar as comissões.
    out = out.drop_duplicates().reset_index(drop=True)
    return out


# Chave de merge: substitui todo o período (preserva o histórico de outros meses).
UNION_KEYS = ["plataforma", "operador", "username", "month"]

# Renomeação/ordenação final das colunas (interno -> nome de saída).
# 'Origen' e 'Automatizado' são derivados/constantes (ver build_final()).
FINAL_RENAME = {
    "index": "Index",
    "month": "AnoMes",
    "empresa": "Marca",
    "plataforma": "Plataforma",
    "ref": "ReferenciaFacturacion",
    "operador": "Operador",
    "username": "Login",
    "resta": "Resta",
    "currency": "Moneda_Operador",
    "rate": "ValorMoneda_XE",
    "rs_operador": "RS_Operador",
    "rs_eur": "RS_EUR",
    "cpa_operador": "CPA_Operador",
    "cpa_eur": "CPA_EUR",
}
FINAL_COLS = [
    "Index", "AnoMes", "Marca", "Plataforma", "Automatizado",
    "ReferenciaFacturacion", "Origen", "Operador", "Login", "Resta",
    "Moneda_Operador", "ValorMoneda_XE", "RS_Operador", "RS_EUR",
    "CPA_Operador", "CPA_EUR",
]
# Chave de merge sobre as colunas já renomeadas.
FINAL_KEYS = ["Plataforma", "Operador", "Login", "AnoMes"]


def total_facturacion(resta, cpa, rs) -> float:
    """Total de facturação segundo a regra do negócio (erro/não-numérico -> 0)."""
    try:
        cpa = float(cpa)
        rs = float(rs)
        r = str(resta).strip()
        if r == "No":
            return round(cpa + (0 if rs < 0 else rs), 2)
        elif r == "Sí":
            if cpa + rs < 0:
                return 0.0
            if rs < 0 and cpa == 0:
                return round(cpa + 0, 2)
            return round(cpa + rs, 2)
        else:
            return round(cpa + rs, 2)
    except Exception:
        return 0.0


def normalize_login(username: str) -> str:
    """Login normalizado. Para e-mails, junta a parte local ao nome do domínio
    (sem TLD), ignorando 'gmail'. Ex.: 'info@afiliagambling.com' -> 'infoafiliagambling';
    'affiliabetbrasil@gmail.com' -> 'affiliabetbrasil'. Não-e-mail fica inalterado."""
    s = str(username).strip()
    if "@" not in s:
        return s
    local, _, domain = s.partition("@")
    sld = domain.split(".")[0]
    if sld.lower() == "gmail":
        return local
    return f"{local}{sld}"


def _upsert_csv(
    path: Path,
    new_df: pd.DataFrame,
    key_cols: list[str],
    replace_keys: pd.DataFrame | None = None,
) -> pd.DataFrame:
    return common.upsert_csv(path, new_df, key_cols, replace_keys=replace_keys, log=log)


def _load_rates() -> dict[tuple[str, str], float]:
    """Lê o currency_rates.csv -> {(month, currency): rate} (base EUR: 1 EUR = rate).

    Quando há várias cotações para o mesmo (month, currency), usa a MAIS ANTIGA do mês
    (menor time_last_update_utc), por ser a mais próxima do fecho do período.
    """
    if not RATES_CSV.exists():
        log.warning(f"{RATES_CSV.name} não encontrado — rs_eur/cpa_eur ficam vazios")
        return {}
    df = pd.read_csv(RATES_CSV, dtype=str)
    # Ordena pela data/hora da cotação (ascendente) e fica com a mais antiga por chave.
    df["_ts"] = pd.to_datetime(df["time_last_update_utc"], errors="coerce")
    df = df.sort_values("_ts", na_position="last")
    df = df.drop_duplicates(subset=["month", "currency"], keep="first")
    rates: dict[tuple[str, str], float] = {}
    for _, r in df.iterrows():
        try:
            rates[(str(r["month"]).strip(), str(r["currency"]).strip())] = float(r["rate"])
        except (ValueError, KeyError):
            continue
    return rates


def _to_eur(value: float, month: str, currency: str, rates: dict) -> float | None:
    """Converte um valor na 'currency' para EUR usando a taxa do mês (valor/rate)."""
    rate = rates.get((str(month).strip(), str(currency).strip()))
    if rate is None or rate == 0:
        return None
    return round(float(value) / rate, 2)


def enrich_ref(union: pd.DataFrame) -> pd.DataFrame:
    """Junta as colunas 'ref' e 'resta' do config/ref.xlsx.

    Chave normal: (operador, username, empresa). Excepção: para Novibet na empresa
    Afiliagambling, o 'operador' do ref guarda o nome do vendor, por isso a junção
    usa o 'vendor_name' no lugar do operador.
    """
    union = union.copy()
    if not REF_XLSX.exists():
        log.warning(f"{REF_XLSX.name} não encontrado — ref/resta ficam vazios")
        union["ref"] = ""
        union["resta"] = ""
        return union

    def norm(s):  # chave case-insensitive (resolve diferenças de caixa entre fontes)
        return s.fillna("").astype(str).str.strip().str.lower()

    ref = pd.read_excel(REF_XLSX, dtype=str)
    ref = ref[["operador", "username", "empresa", "ref", "resta"]].copy()
    ref["_jk"] = norm(ref["operador"])
    ref["_un"] = norm(ref["username"])
    ref["_em"] = norm(ref["empresa"])
    ref = ref.drop_duplicates(subset=["_jk", "_un", "_em"], keep="first")
    ref = ref[["_jk", "_un", "_em", "ref", "resta"]]

    op = norm(union["operador"])
    emp = norm(union["empresa"])
    ven = norm(union["vendor_name"])
    # Excepção Novibet/Afiliagambling: junta pelo vendor_name.
    excecao = (op == "novibet") & (emp == "afiliagambling")
    union["_jk"] = op.where(~excecao, ven)
    union["_un"] = norm(union["username"])
    union["_em"] = emp

    merged = union.merge(ref, on=["_jk", "_un", "_em"], how="left")
    sem_ref = merged["ref"].isna().sum()
    if sem_ref:
        log.warning(f"{sem_ref} linha(s) sem correspondência em ref.xlsx")
    return merged.drop(columns=["_jk", "_un", "_em"])


def _accounts_for_ids(ids: list[str]) -> list[tuple[str, str]]:
    """Devolve lista de (operador_lower, username_lower) para os IDs fornecidos."""
    logins_path = ROOT / "config" / "logins.xlsx"
    if not logins_path.exists():
        raise FileNotFoundError(f"{logins_path} não encontrado")
    lg = pd.read_excel(logins_path, dtype=str)
    wanted = {str(i).strip() for i in ids}
    rows = lg[lg["Id"].astype(str).str.strip().isin(wanted)]
    if rows.empty:
        raise ValueError(f"Nenhum ID encontrado no logins.xlsx: {ids}")
    pairs = []
    for _, r in rows.iterrows():
        op = str(r["Operador"]).strip().lower()
        un = str(r["Username"]).strip().lower()
        log.info(f"ID {r['Id'].strip()} → {r['Operador']} / {r['Username']} ({r['Plataforma']} / {r['Empresa']})")
        pairs.append((op, un))
    return pairs


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="Constrói report/union_data.csv")
    parser.add_argument(
        "--id", nargs="+", dest="ids",
        help="Processa e grava apenas as contas dos IDs indicados (ex: --id 148 149)",
    )
    args = parser.parse_args()

    log.info("A construir report/union_data.csv ...")
    frames = [
        from_netrefer(),
        from_raventrack(),
        from_cellxpert(),
        from_income_access(),
    ]
    union = pd.concat(frames, ignore_index=True)

    # Corrige nomes de operador renomeados no login (ex.: Sportium -> SportiumBet).
    union["operador"] = union["operador"].map(common.canonical_operador)

    union["rs_operador"] = union["rs_operador"].round(2)
    union["cpa_operador"] = union["cpa_operador"].round(2)

    # Preenche a moeda APENAS quando está vazia/NaN: 'Brasil' -> BRL, restante -> EUR.
    # Se a coluna currency já tiver valor, mantém o que tem.
    cur = union["currency"].astype(str).str.strip()
    # "0" aparece no income_access como placeholder de contas sem dados (não é moeda).
    vazio = cur.isin(["", "0", "nan", "NaN", "None"]) | union["currency"].isna()
    empresa = union["empresa"].fillna("").astype(str).str.strip().str.lower()
    fallback = empresa.map(lambda e: "BRL" if e == "brasil" else "EUR")
    union["currency"] = cur.where(~vazio, fallback)

    # SportiumBet em USD corresponde à conta do Panamá: ajusta o username para
    # 'Affiliabet_Panama' (antes do join), que casa com o ref Panamá no ref.xlsx.
    sport_usd = (union["operador"].astype(str).str.strip() == "SportiumBet") & (
        union["currency"].astype(str).str.strip() == "USD"
    )
    union.loc[sport_usd, "username"] = "Affiliabet_Panama"

    # Converte rs/cpa para EUR usando o câmbio do mês (join por month + currency).
    rates = _load_rates()
    sem_taxa = set()
    rs_eur, cpa_eur = [], []
    for _, row in union.iterrows():
        m, c = row["month"], row["currency"]
        if (str(m).strip(), str(c).strip()) not in rates and c:
            sem_taxa.add((str(m).strip(), str(c).strip()))
        rs_eur.append(_to_eur(row["rs_operador"], m, c, rates))
        cpa_eur.append(_to_eur(row["cpa_operador"], m, c, rates))
    union["rs_eur"] = rs_eur
    union["cpa_eur"] = cpa_eur
    if sem_taxa:
        log.warning(f"sem taxa de câmbio para {sorted(sem_taxa)} — rs_eur/cpa_eur vazios nessas linhas")

    # Empresa Brasil reportada em EUR: os valores '_operador' devem ficar em BRL
    # (= valor EUR * taxa BRL do mês), mantendo '_eur' em EUR e currency = BRL.
    emp = union["empresa"].fillna("").astype(str).str.strip().str.lower()
    br_eur = (emp == "brasil") & (union["currency"].astype(str).str.strip() == "EUR")
    convertidas = 0
    for idx in union.index[br_eur]:
        m = union.at[idx, "month"]
        rate_brl = rates.get((str(m).strip(), "BRL"))
        if rate_brl is None or rate_brl == 0:
            continue
        if union.at[idx, "rs_eur"] is not None:
            union.at[idx, "rs_operador"] = round(float(union.at[idx, "rs_eur"]) * rate_brl, 2)
        if union.at[idx, "cpa_eur"] is not None:
            union.at[idx, "cpa_operador"] = round(float(union.at[idx, "cpa_eur"]) * rate_brl, 2)
        union.at[idx, "currency"] = "BRL"
        convertidas += 1
    if convertidas:
        log.info(f"{convertidas} linha(s) Brasil/EUR convertidas: _operador em BRL, _eur em EUR")

    # Taxa de câmbio do mês para a moeda final de cada linha (base EUR: 1 EUR = rate).
    union["rate"] = [
        rates.get((str(m).strip(), str(c).strip()))
        for m, c in zip(union["month"], union["currency"])
    ]

    # Enriquece com ref/resta do config/ref.xlsx.
    union = enrich_ref(union)

    # Filtro por ID: mantém apenas as contas dos IDs pedidos.
    if args.ids:
        pairs = _accounts_for_ids(args.ids)
        op_col = union["operador"].astype(str).str.strip().str.lower()
        un_col = union["username"].astype(str).str.strip().str.lower()
        mask = pd.Series(False, index=union.index)
        for op, un in pairs:
            mask |= (op_col == op) & (un_col == un)
        before = len(union)
        union = union[mask].reset_index(drop=True)
        log.info(f"Filtro --id: {len(union)} linha(s) de {before} seleccionadas")

    # Chaves de todos os períodos processados (antes de remover), para limpar também
    # do histórico as linhas que agora foram removidas (inactivas/agregadas).
    processed_keys = union[UNION_KEYS].copy()

    # O ref.xlsx é a lista-mestre do que deve aparecer. Linhas sem correspondência
    # são contas inactivas ou agregados (ex.: a conta-pai TipsterpageAR, cujos dados
    # válidos são apenas as sub-contas PBA_/CABA_) e são removidas.
    sem_ref = union["ref"].isna()
    removidas = union[sem_ref][["plataforma", "operador", "username", "empresa"]]
    # Relatório de divergências de cadastro: reescrito SEMPRE (vazio quando não há
    # divergência), para não persistir um falso-positivo de uma execução anterior.
    div_path = REPORT_DIR / "divergencias_cadastro.csv"
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    removidas.to_csv(div_path, index=False, encoding="utf-8-sig")
    if sem_ref.any():
        log.info(f"{int(sem_ref.sum())} linha(s) removidas (sem ref / inactivas):")
        for _, r in removidas.iterrows():
            log.info(f"    - {r['plataforma']} | {r['operador']} | {r['username']} | {r['empresa']}")
        log.info(f"Divergências de cadastro: {div_path}")
        union = union[~sem_ref].reset_index(drop=True)

    # 'resta' vazio no ref.xlsx assume o valor 'No'.
    resta = union["resta"].astype(str).str.strip()
    union["resta"] = resta.where(~resta.isin(["", "nan", "NaN", "None"]) & union["resta"].notna(), "No")

    # Coluna 'index' = concatenação de month, ref, plataforma e username (sem separadores).
    union["index"] = (
        union["month"].astype(str).str.strip()
        + union["ref"].astype(str).str.strip()
        + union["plataforma"].astype(str).str.strip()
        + union["username"].astype(str).str.strip()
    )

    union = union[OUTPUT_COLS]

    # Origen = Operador_Login (login normalizado p/ e-mails) e Automatizado constante.
    origen = (
        union["operador"].astype(str).str.strip()
        + "_" + union["username"].map(normalize_login)
    )
    union["Automatizado"] = "Sí"
    union["Origen"] = origen

    # Renomeia e reordena para o esquema final.
    union = union.rename(columns=FINAL_RENAME)[FINAL_COLS]

    # Total de facturação (regra de negócio); _EUR usa as colunas em EUR.
    union["TotalFacturacion_Operador"] = [
        total_facturacion(r, c, s)
        for r, c, s in zip(union["Resta"], union["CPA_Operador"], union["RS_Operador"])
    ]
    union["TotalFacturacion_EUR"] = [
        total_facturacion(r, c, s)
        for r, c, s in zip(union["Resta"], union["CPA_EUR"], union["RS_EUR"])
    ]

    # Excepção RavenTrack/affiliabetbrasil@gmail.com: os campos '_Operador' ficam
    # com o texto "Esperar cambio" (câmbio ainda não disponível); os '_EUR' mantêm
    # os valores normais.
    esperar_cambio = (
        (union["Plataforma"].astype(str).str.strip() == "Raventrack")
        & (union["Login"].astype(str).str.strip().str.lower() == "affiliabetbrasil@gmail.com")
    )
    if esperar_cambio.any():
        for col in ["RS_Operador", "CPA_Operador", "TotalFacturacion_Operador"]:
            union[col] = union[col].astype(object)
            union.loc[esperar_cambio, col] = "Esperar cambio"
        # ValorMoneda_XE fica vazio enquanto o câmbio não for definido na planilha
        # (é o gatilho que a fórmula do Sheets usa para saber quando calcular).
        union["ValorMoneda_XE"] = union["ValorMoneda_XE"].astype(object)
        union.loc[esperar_cambio, "ValorMoneda_XE"] = ""

    processed_keys = processed_keys.rename(
        columns={"plataforma": "Plataforma", "operador": "Operador",
                 "username": "Login", "month": "AnoMes"}
    )

    out_path = REPORT_DIR / "union_data.csv"
    # Merge com o histórico já guardado: substitui só os períodos do lote novo
    # (incluindo as chaves removidas, para que saiam também do histórico).
    merged = _upsert_csv(
        out_path, union.astype(str), FINAL_KEYS, replace_keys=processed_keys.astype(str)
    )
    # Mantém apenas o mês alvo no ficheiro final.
    target_month = common.previous_month_label()
    merged = merged[merged["AnoMes"] == target_month].reset_index(drop=True)
    merged.to_csv(out_path, index=False, encoding="utf-8-sig")
    log.info(f"OK -> {out_path}  ({len(union)} novas, {len(merged)} no total)")
    log.info("Linhas por plataforma/mês:\n" + merged.groupby(["Plataforma", "AnoMes"]).size().to_string())


if __name__ == "__main__":
    main()
