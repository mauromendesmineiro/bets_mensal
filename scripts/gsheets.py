"""
scripts/gsheets.py

Integração com o Google Sheets (leitura/escrita) via Service Account.

Funções:
- ref_from_sheet()        — lê as 4 abas e actualiza a coluna 'resta' do
                            config/ref.xlsx (correspondência: empresa + username +
                            ref; 'resta' vazio assume 'No'). Lê a planilha
                            directamente, sem ficheiro temporário.
- write_union_to_sheet()  — escreve o report/union_data.csv numa aba da planilha.

Autenticação por OAuth (conta do próprio utilizador, que já tem acesso à planilha):
1. No Google Cloud, com a Google Sheets API activada, cria credenciais OAuth do tipo
   "App de computador" (Desktop) e descarrega o JSON.
2. Guarda-o em  config/oauth_client.json  (NÃO comitar).
3. Na 1ª execução abre o navegador para autorizares com a tua conta Google; o token
   fica salvo em  config/authorized_user.json  para as próximas vezes.

Configuração via .env:
    GOOGLE_OAUTH_CLIENT_JSON=config/oauth_client.json       # credenciais OAuth (Desktop)
    GOOGLE_OAUTH_TOKEN_JSON=config/authorized_user.json     # token gerado no 1º login
    GSHEET_ID=<id da planilha>            # o trecho entre /d/ e /edit no URL
    GSHEET_REF_TAB=ref                    # nome da aba de referência
    GSHEET_UNION_TAB=union_data           # nome da aba de destino do union

Uso:
    python scripts/gsheets.py --pull-ref     # planilha -> config/ref.xlsx (merge)
    python scripts/gsheets.py --push-union    # union_data.csv -> planilha
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd

try:  # funciona tanto em execução directa quanto importado como scripts.*
    import common
except ImportError:
    from scripts import common

common.load_env()

log = common.get_logger("gsheets")

ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = ROOT / "config"
REPORT_DIR = ROOT / "report"
REF_XLSX = CONFIG_DIR / "ref.xlsx"
LOGINS_XLSX = CONFIG_DIR / "logins.xlsx"
MISSING_CSV = CONFIG_DIR / "contas_faltantes.csv"
UNION_CSV = REPORT_DIR / "union_data.csv"

# Merge planilha -> ref.xlsx. A correspondência usa estas colunas:
#   ref.xlsx 'empresa'  <-> planilha 'empresa'
#   ref.xlsx 'username' <-> planilha 'Usuario Referencia'
#   ref.xlsx 'ref'      <-> planilha 'Operador Referencia'
# A coluna actualizada é 'resta' (planilha 'Resta'); vazio assume 'No'.
REF_COLS = ["operador", "username", "empresa", "ref", "resta"]
REF_MATCH = [("empresa", "empresa"), ("username", "Usuario Referencia"), ("ref", "Operador Referencia")]
RESTA_SHEET_COL = "Resta"
RESTA_DEFAULT = "No"

# As 4 abas do ref na planilha → empresa (derivada do nome da aba).
REF_TABS = {
    "RS/CPA_Affiliabet": "Affiliabet",
    "RS/CPA_Afiliagambling": "Afiliagambling",
    "RS/CPA_Afiliawin": "Afiliawin",
    "RS/CPA_Brasil": "Brasil",
}
# Linha do cabeçalho (1-based) e colunas necessárias (lidas por nome).
REF_HEADER_ROW = 9
REF_SHEET_COLS = [
    "Plataforma", "Usuario Referencia",
    "Operador Referencia", "RS / CPA / Fijo", "Resta",
]
# Só interessam os registos de comissão RS ou CPA (outros são removidos).
# Preferência por chave: RS primeiro; se não houver, CPA.
REF_COMMISSION_COL = "RS / CPA / Fijo"
REF_COMMISSION_KEEP = {"rs", "cpa"}
REF_COMMISSION_RANK = {"rs": 0, "cpa": 1}
REF_DEDUP_KEY = ["empresa", "Usuario Referencia", "Operador Referencia"]
# Só interessam as linhas das plataformas que extraímos (comparação case-insensitive).
USED_PLATFORMS = common.USED_PLATFORMS


# ── Configuração ────────────────────────────────────────────────────────────
def _cfg(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def _abs(path_str: str) -> Path:
    p = Path(path_str)
    return p if p.is_absolute() else ROOT / p


# ── Cliente gspread (OAuth com a conta do utilizador) ───────────────────────
def _open_spreadsheet():
    """Abre a planilha definida em GSHEET_ID via OAuth (conta do utilizador)."""
    import gspread  # import tardio: só é necessário quando se usa o Sheets

    client_path = _abs(_cfg("GOOGLE_OAUTH_CLIENT_JSON", "config/oauth_client.json"))
    token_path = _abs(_cfg("GOOGLE_OAUTH_TOKEN_JSON", "config/authorized_user.json"))
    if not client_path.exists():
        raise FileNotFoundError(
            f"Credenciais OAuth não encontradas: {client_path}\n"
            "Cria credenciais OAuth (App de computador) no Google Cloud e guarda em "
            "config/oauth_client.json."
        )

    sheet_id = _cfg("GSHEET_ID")
    if not sheet_id:
        raise ValueError("GSHEET_ID não definido no .env")

    # 1ª vez: abre o navegador para autorizar; depois reutiliza o token salvo.
    gc = gspread.oauth(
        credentials_filename=str(client_path),
        authorized_user_filename=str(token_path),
    )
    return gc.open_by_key(sheet_id)


def _read_tab_with_header(sh, tab: str, header_row: int, cols: list[str]) -> pd.DataFrame:
    """Lê uma aba cujo cabeçalho está na linha `header_row`, devolvendo só `cols`."""
    values = sh.worksheet(tab).get_all_values()
    if len(values) < header_row:
        return pd.DataFrame(columns=cols)
    header = [h.strip() for h in values[header_row - 1]]
    data = values[header_row:]  # linhas após o cabeçalho
    df = pd.DataFrame(data, columns=header)
    # Selecciona apenas as colunas pedidas (as ausentes ficam vazias).
    out = pd.DataFrame()
    for c in cols:
        out[c] = df[c] if c in df.columns else ""
    # Remove linhas totalmente vazias nas colunas pedidas.
    out = out[~(out.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    return out.reset_index(drop=True)


def read_ref_sheets() -> pd.DataFrame:
    """Lê e combina as 4 abas do ref. Acrescenta 'empresa' (do nome da aba)."""
    sh = _open_spreadsheet()
    frames = []
    for tab, empresa in REF_TABS.items():
        df = _read_tab_with_header(sh, tab, REF_HEADER_ROW, REF_SHEET_COLS)
        df.insert(0, "empresa", empresa)
        frames.append(df)
        log.info(f"aba '{tab}': {len(df)} linha(s)")
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    # Mantém só as plataformas que extraímos.
    plat = combined["Plataforma"].fillna("").astype(str).str.strip().str.lower()
    combined = combined[plat.isin(USED_PLATFORMS)].reset_index(drop=True)
    # Mantém só RS ou CPA (remove Fijo/Fee/vazio/etc.).
    com = combined[REF_COMMISSION_COL].fillna("").astype(str).str.strip().str.lower()
    combined = combined[com.isin(REF_COMMISSION_KEEP)].reset_index(drop=True)
    # Preferência por chave: RS antes de CPA — mantém 1 linha por conta.
    com = combined[REF_COMMISSION_COL].astype(str).str.strip().str.lower()
    combined["_rank"] = com.map(REF_COMMISSION_RANK).fillna(9)
    combined = (
        combined.sort_values("_rank")
        .drop_duplicates(subset=REF_DEDUP_KEY, keep="first")
        .drop(columns="_rank")
        .sort_index()
        .reset_index(drop=True)
    )
    return combined


# ── 1) Planilha -> ref.xlsx (merge da coluna 'resta') ───────────────────────
def _norm(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.lower()


def ref_from_sheet() -> Path:
    """Lê as 4 abas da planilha e actualiza a coluna 'resta' do config/ref.xlsx,
    correspondendo por empresa + username + ref (ver REF_MATCH). 'resta' vazio
    assume 'No'. Lê a planilha directamente (sem ficheiro temporário)."""
    sheet = read_ref_sheets()
    if sheet.empty:
        log.warning("nenhuma linha lida das abas — ref.xlsx não alterado")
        return REF_XLSX
    if not REF_XLSX.exists():
        raise FileNotFoundError(f"{REF_XLSX} não encontrado")

    ref = pd.read_excel(REF_XLSX, dtype=str)

    # Lookup: chave (empresa, username, ref) -> Resta da planilha.
    sheet_keys = [_norm(sheet[s]) for _, s in REF_MATCH]
    sheet_lookup = {
        k: r for k, r in zip(zip(*sheet_keys), sheet[RESTA_SHEET_COL].fillna(""))
    }
    ref_keys = list(zip(*[_norm(ref[r]) for r, _ in REF_MATCH]))

    atualizadas = 0
    novo_resta = []
    for k, atual in zip(ref_keys, ref["resta"]):
        if k in sheet_lookup:
            val = str(sheet_lookup[k]).strip()
            atualizadas += 1
        else:
            val = str(atual).strip()
        if val.lower() in ("", "nan", "none"):
            val = RESTA_DEFAULT
        novo_resta.append(val)
    ref["resta"] = novo_resta

    ref.to_excel(REF_XLSX, index=False)
    log.info(f"OK -> {REF_XLSX}  ({len(ref)} linhas; resta actualizada em {atualizadas})")
    return REF_XLSX


# ── 3) Validação: contas da planilha em falta no logins.xlsx ────────────────
def _plat_norm(s: pd.Series) -> pd.Series:
    """Normaliza a plataforma para comparar fontes distintas.

    A planilha escreve 'Income Access' e o logins.xlsx 'IncomeAccess'; remover
    espaços + minúsculas unifica ambos ('incomeaccess')."""
    return s.fillna("").astype(str).str.strip().str.lower().str.replace(" ", "", regex=False)


def _is_base_login(sheet_user: str, login_user: str) -> bool:
    """True se `login_user` for a login-base de `sheet_user`.

    Cobre os sub-registos sintéticos do ref/planilha que partilham uma login
    real, separados por moeda/região por prefixo ou sufixo delimitado por '_':
    'affiliabet_panama' ← 'affiliabet', 'caba_tipsterpagear' ← 'tipsterpagear'.
    """
    if not login_user:
        return False
    return (
        sheet_user == login_user
        or sheet_user.startswith(login_user + "_")
        or sheet_user.endswith("_" + login_user)
        or ("_" + login_user + "_") in sheet_user
    )


def validate_logins() -> Path:
    """Compara as contas da planilha (4 abas) com o config/logins.xlsx e grava
    config/contas_faltantes.csv com as que existem na planilha mas não no logins.

    A correspondência é por (empresa, plataforma, username). Sub-registos
    sintéticos (split por moeda/região, ex.: '..._panama', 'caba_...') são
    considerados presentes quando a sua login-base existe no logins — evita
    falsos-positivos. Inserir contas novas no logins.xlsx/ref.xlsx é manual.
    """
    if not LOGINS_XLSX.exists():
        raise FileNotFoundError(f"{LOGINS_XLSX} não encontrado")
    sheet = read_ref_sheets()
    if sheet.empty:
        log.warning("nenhuma linha lida das abas — validação não executada")
        return MISSING_CSV

    lg = pd.read_excel(LOGINS_XLSX, dtype=str)
    lg_emp = _norm(lg["Empresa"])
    lg_plat = _plat_norm(lg["Plataforma"])
    lg_user = _norm(lg["Username"])
    # Índice (empresa, plataforma) -> conjunto de usernames de login.
    by_scope: dict[tuple[str, str], set[str]] = {}
    for e, p, u in zip(lg_emp, lg_plat, lg_user):
        by_scope.setdefault((e, p), set()).add(u)

    s_emp = _norm(sheet["empresa"])
    s_plat = _plat_norm(sheet["Plataforma"])
    s_user = _norm(sheet["Usuario Referencia"])

    faltam_idx = []
    for i, (e, p, u) in enumerate(zip(s_emp, s_plat, s_user)):
        logins = by_scope.get((e, p), set())
        if u in logins or any(_is_base_login(u, lu) for lu in logins):
            continue
        faltam_idx.append(i)

    cols = ["empresa", "Plataforma", "Usuario Referencia", "Operador Referencia",
            "RS / CPA / Fijo", "Resta"]
    faltam = sheet.iloc[faltam_idx][cols].reset_index(drop=True)
    faltam.columns = ["empresa", "plataforma", "usuario_referencia",
                      "operador_referencia", "comissao", "resta"]
    faltam.to_csv(MISSING_CSV, index=False, encoding="utf-8-sig")

    if faltam.empty:
        log.info(f"OK -> todas as contas da planilha estão no logins.xlsx ({MISSING_CSV.name} vazio)")
    else:
        log.warning(f"{len(faltam)} conta(s) na planilha SEM login configurado:")
        for _, r in faltam.iterrows():
            log.warning(f"    - {r['empresa']} | {r['plataforma']} | "
                        f"{r['usuario_referencia']} | {r['operador_referencia']}")
        log.info(f"-> {MISSING_CSV}")
    return MISSING_CSV


# ── 2) union_data.csv -> planilha ───────────────────────────────────────────
def write_union_to_sheet(csv_path: Path = UNION_CSV) -> None:
    """Escreve o conteúdo do union_data.csv na aba GSHEET_UNION_TAB (substitui tudo)."""
    if not csv_path.exists():
        raise FileNotFoundError(f"{csv_path} não encontrado — corra o build_union primeiro")

    df = pd.read_csv(csv_path, dtype=str, encoding="utf-8-sig").fillna("")
    tab = _cfg("GSHEET_UNION_TAB", "union_data")
    sh = _open_spreadsheet()

    try:
        ws = sh.worksheet(tab)
    except Exception:
        ws = sh.add_worksheet(title=tab, rows=len(df) + 10, cols=len(df.columns) + 2)

    ws.clear()
    ws.update([df.columns.tolist()] + df.values.tolist(), value_input_option="RAW")
    log.info(f"OK -> aba '{tab}' actualizada ({len(df)} linhas, {len(df.columns)} colunas)")


def _convert_to_comma_decimal(value: str) -> str:
    """Converte separadores de decimal de ponto para vírgula em valores numéricos.

    Se o valor parece numérico (tem ponto e dígitos), substitui ponto por vírgula.
    Ex.: '123.45' → '123,45'; 'texto' → 'texto' (inalterado).
    """
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return ""
    # Se tem ponto e só tem dígitos/ponto/hífen/espaço, é provavelmente numérico
    if "." in s and all(c.isdigit() or c in ".-+ " for c in s):
        return s.replace(".", ",")
    return s


def append_union_to_sheet(csv_path: Path = UNION_CSV) -> None:
    """Escreve union_data.csv na aba DatosAutomatizados (append: adiciona ao final).

    Converte valores numéricos para vírgula como separador de decimal.
    Se a aba não existe, cria; se existe, acrescenta após a última linha de dados.
    """
    if not csv_path.exists():
        raise FileNotFoundError(f"{csv_path} não encontrado — corra o build_union primeiro")

    df = pd.read_csv(csv_path, dtype=str, encoding="utf-8-sig").fillna("")
    tab = "DatosAutomatizados"
    sh = _open_spreadsheet()

    try:
        ws = sh.worksheet(tab)
    except Exception:
        ws = sh.add_worksheet(title=tab, rows=len(df) + 100, cols=len(df.columns) + 2)
        # Primeira vez: adiciona cabeçalho
        ws.append_row(df.columns.tolist(), value_input_option="USER_ENTERED")
        log.debug(f"Aba '{tab}' criada com cabeçalho")

    # Converte decimais em todos os valores
    df_converted = df.map(lambda x: _convert_to_comma_decimal(x))

    # Append em bloco: envia todas as linhas de uma vez (economiza quota do Sheets)
    # USER_ENTERED permite que o Sheets interprete números, datas, etc. corretamente
    rows_to_append = df_converted.values.tolist()
    ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")

    log.info(f"OK -> aba '{tab}' actualizada (append: {len(df)} linhas adicionadas)")


# ── CLI ─────────────────────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="Integração Google Sheets (ref / union)")
    parser.add_argument("--pull-ref", action="store_true", help="Planilha (4 abas) -> merge 'resta' no config/ref.xlsx")
    parser.add_argument("--push-union", action="store_true", help="union_data.csv -> planilha (substitui tudo)")
    parser.add_argument("--append-union", action="store_true", help="union_data.csv -> planilha aba DatosAutomatizados (append, com decimais em vírgula)")
    parser.add_argument("--validate-logins", action="store_true",
                        help="Planilha vs logins.xlsx -> config/contas_faltantes.csv")
    args = parser.parse_args()

    if not (args.pull_ref or args.push_union or args.append_union or args.validate_logins):
        parser.print_help()
        sys.exit(0)

    if args.pull_ref:
        ref_from_sheet()
    if args.validate_logins:
        validate_logins()
    if args.push_union:
        write_union_to_sheet()
    if args.append_union:
        append_union_to_sheet()


if __name__ == "__main__":
    main()
