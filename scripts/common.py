"""
scripts/common.py

Utilitários partilhados pelos scrapers e pelo pipeline (item 1 do MELHORIAS.md).
Centraliza o que estava duplicado em scp_netrefer/income_access/cellxpert/raventrack
e em build_union: carregamento do .env, logger, normalização, chave de password,
rótulo do mês anterior, upsert de CSV, helpers de Playwright e detecção de moeda.

Migração incremental: cada módulo pode passar a importar daqui sem alterar o seu
comportamento (as funções são equivalentes às cópias locais).
"""

from __future__ import annotations

import logging
import os
import re
import sys
import time
from collections.abc import Callable
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd

# ── Caminhos do projecto ────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parents[1]
CONFIG_XLSX = ROOT / "config" / "logins.xlsx"
DATA_DIR = ROOT / "data"
REPORT_DIR = ROOT / "report"
LOGS_DIR = ROOT / "logs"


# ── .env ────────────────────────────────────────────────────────────────────
def load_env() -> None:
    """Carrega o .env da raiz suportando chaves entre aspas ("KEY"="VALUE")."""
    env_file = ROOT / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        k = k.strip().strip('"')
        v = v.strip().strip('"')
        if k and k not in os.environ:
            os.environ[k] = v


# Carrega o .env já no import do common, para que os settings abaixo (lidos do
# ambiente) reflictam o .env. Os scrapers podem chamar load_env() de novo (idempotente).
load_env()


# ── Settings (lidos do ambiente / .env) ─────────────────────────────────────
DEFAULT_TIMEOUT = int(os.getenv("DEFAULT_TIMEOUT", "20000"))  # timeout Playwright (ms)
SLOW_MO = int(os.getenv("SLOW_MO", "0"))                       # ms entre acções (debug)
HEADLESS = os.getenv("HEADLESS", "true").lower() == "true"
# Tempo máx. à espera de um report pesado terminar de gerar (Income Access) (ms).
REPORT_READY_TIMEOUT = int(os.getenv("REPORT_READY_TIMEOUT", "300000"))
# Retry por conta (falhas transitórias): nº de tentativas e backoff entre elas.
RETRY_ATTEMPTS = int(os.getenv("RETRY_ATTEMPTS", "3"))
RETRY_BACKOFF_S = float(os.getenv("RETRY_BACKOFF_S", "10"))

# ── Constantes de domínio ───────────────────────────────────────────────────
# Plataformas que extraímos (nomes como aparecem na coluna Plataforma da planilha ref).
USED_PLATFORMS = {"netrefer", "income access", "cellxpert", "raventrack"}
# Tipos de comissão garantidos no output Cellxpert mesmo sem dados (amount 0).
DEFAULT_COMMISSION_TYPES = ["CPA", "Revshare Ongoing PL"]

# Aliases de operador renomeados na origem (chave normalizada -> nome correcto).
# Ex.: 'Sportium' foi renomeado para 'SportiumBet' no login.
OPERADOR_ALIASES = {"sportium": "SportiumBet"}


def norm_key(val) -> str:
    """Normaliza um valor para comparação/junção case-insensitive (strip + lower)."""
    if val is None:
        return ""
    return str(val).strip().lower()


def canonical_operador(operador) -> str:
    """Aplica o alias canónico ao operador (ver OPERADOR_ALIASES); senão devolve igual."""
    o = "" if operador is None else str(operador)
    return OPERADOR_ALIASES.get(o.strip().lower(), o)


# ── Logger ──────────────────────────────────────────────────────────────────
def get_logger(name: str) -> logging.Logger:
    """Logger com saída no stdout (INFO) e ficheiro diário por plataforma (DEBUG).

    O ficheiro é logs/<data>_<plataforma>.log, onde <plataforma> deriva do nome
    (ex.: 'scp_cellxpert' -> 'cellxpert')."""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)-7s] %(message)s", datefmt="%H:%M:%S")
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.INFO)
    sh.setFormatter(fmt)
    logger.addHandler(sh)
    try:
        LOGS_DIR.mkdir(parents=True, exist_ok=True)
        plat = name.replace("scp_", "")
        fh = logging.FileHandler(LOGS_DIR / f"{datetime.now():%Y-%m-%d}_{plat}.log", encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        pass
    return logger


# ── Normalização / credenciais ──────────────────────────────────────────────
def sanitize(val: str) -> str:
    """Normaliza um valor para compor nomes de variáveis de ambiente."""
    return re.sub(r"[^A-Z0-9]", "_", str(val).upper())


def env_key_for(slug: str, operador: str, username: str) -> str:
    """Convenção da password no .env: PASS_<SLUG>_<OPERADOR>_<USERNAME>."""
    return f"PASS_{slug}_{sanitize(operador)}_{sanitize(username)}"


# ── Datas ───────────────────────────────────────────────────────────────────
def previous_month_label(today: date | None = None) -> str:
    """Rótulo YYYY-MM do período a extrair.

    Por omissão, o mês anterior a hoje. Pode ser sobreposto pela variável de
    ambiente TARGET_MONTH (formato YYYY-MM) para reprocessar um mês específico —
    só quando `today` não é passado explicitamente (mantém os testes determinísticos).
    """
    if today is None:
        override = os.getenv("TARGET_MONTH", "").strip()
        if re.fullmatch(r"\d{4}-\d{2}", override):
            return override
        today = date.today()
    first = today.replace(day=1)
    prev = (first - timedelta(days=1)).replace(day=1)
    return prev.strftime("%Y-%m")


def safe_name(name: str) -> str:
    """Nome de ficheiro seguro a partir de um texto arbitrário."""
    return re.sub(r"[^A-Za-z0-9_.-]", "_", name).strip("_") or "report"


# ── Retry ───────────────────────────────────────────────────────────────────
def retry_until(
    fn: Callable[[], object],
    ok: Callable[[object], bool],
    attempts: int = 3,
    backoff_s: float = 10.0,
    log: logging.Logger | None = None,
    label: str = "",
) -> object:
    """Chama fn() até ok(resultado) ser True ou esgotar as tentativas.

    Devolve o último resultado (mesmo que falhe). Entre tentativas espera
    backoff_s segundos. Útil para reprocessar uma conta que falhou por motivo
    transitório (rede/captcha/SPA lenta)."""
    resultado = None
    for i in range(1, max(1, attempts) + 1):
        resultado = fn()
        if ok(resultado):
            if i > 1 and log:
                log.info(f"{label} OK na tentativa {i}/{attempts}")
            return resultado
        if i < attempts:
            if log:
                log.warning(f"{label} tentativa {i}/{attempts} falhou — a repetir em {backoff_s:.0f}s")
            time.sleep(backoff_s)
    return resultado


# ── Contas (logins.xlsx) ────────────────────────────────────────────────────
def load_login_rows(
    plataforma: str,
    operador_filter: str | None = None,
    ids: list[str] | None = None,
) -> pd.DataFrame:
    """Lê o config/logins.xlsx e devolve as linhas activas da plataforma dada.

    - `plataforma`: comparação case-insensitive com a coluna 'Plataforma'.
    - `operador_filter`: opcional, restringe a um operador.
    - `ids`: opcional, restringe à(s) conta(s) pela coluna 'Id' (ver item 12).
    Filtra Active == '1'. Cada scraper constrói o seu Account a partir das linhas.
    """
    if not CONFIG_XLSX.exists():
        return pd.DataFrame()
    df = pd.read_excel(CONFIG_XLSX, dtype=str)
    df = df[
        (df["Plataforma"].fillna("").str.strip().str.lower() == plataforma.strip().lower())
        & (df["Active"].astype(str).str.strip() == "1")
    ].reset_index(drop=True)
    if operador_filter:
        df = df[df["Operador"].fillna("").str.strip().str.lower() == operador_filter.strip().lower()]
    if ids:
        wanted = {str(i).strip() for i in ids}
        df = df[df["Id"].astype(str).str.strip().isin(wanted)]
    return df.reset_index(drop=True)


def base_url_of(login_url: str) -> str:
    p = urlparse(login_url)
    return f"{p.scheme}://{p.netloc}"


# ── Upsert de CSV ───────────────────────────────────────────────────────────
def upsert_csv(
    path: Path,
    new_df: pd.DataFrame,
    key_cols: list[str],
    replace_keys: pd.DataFrame | None = None,
    log: logging.Logger | None = None,
) -> pd.DataFrame:
    """Substitui no CSV existente as linhas cujas key_cols coincidem; mantém o resto.

    `replace_keys`: conjunto de chaves a purgar do histórico (por omissão, as do
    new_df) — útil quando o lote processou chaves depois removidas (inactivas)."""
    keys_src = replace_keys if replace_keys is not None else new_df
    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
            missing = [c for c in key_cols if c not in existing.columns]
            if missing:
                if log:
                    log.warning(f"Upsert: colunas chave ausentes em {path.name}: {missing} — a sobrescrever")
                return new_df.copy()
            mask = existing[key_cols].astype(str).apply(tuple, axis=1).isin(
                keys_src[key_cols].astype(str).apply(tuple, axis=1)
            )
            return pd.concat([existing[~mask], new_df], ignore_index=True)
        except Exception as e:
            if log:
                log.warning(f"Upsert falhou ({e}) — a sobrescrever {path.name}")
    return new_df.copy()


def save_combined(
    frames: list[pd.DataFrame],
    csv_path: Path,
    key_cols: list[str],
    log: logging.Logger,
    as_str: bool = False,
) -> Path | None:
    """Consolida os frames num CSV (upsert por key_cols). Devolve o path ou None.

    `as_str=True` converte tudo para string antes de gravar (cellxpert/raventrack)."""
    if not frames:
        log.warning(f"Nenhum dado para consolidar — {csv_path.name} não gerado")
        return None
    new_data = pd.concat(frames, ignore_index=True)
    if as_str:
        new_data = new_data.astype(str)
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    merged = upsert_csv(csv_path, new_data, key_cols, log=log)
    merged.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log.info(f"Ficheiro consolidado: {csv_path}  ({len(merged)} linhas)")
    return csv_path


def run_accounts(
    accounts: list,
    process_one: Callable[[object], object],
    log: logging.Logger,
) -> tuple[list, list[str], list[str]]:
    """Processa cada conta (com retry por conta) e separa em (frames, sem_dados, erros).

    `process_one(acc)` deve devolver: None (erro), DataFrame vazio (sem dados) ou
    DataFrame com dados. Devolve as três listas para o sumário do scraper."""
    frames: list = []
    sem_dados: list[str] = []
    erros: list[str] = []
    for i, acc in enumerate(accounts, 1):
        label = f"{acc.operador}/{acc.username}"
        log.info(f"[{i}/{len(accounts)}] {acc.operador} / {acc.username}")
        df = retry_until(
            lambda acc=acc: process_one(acc),
            ok=lambda r: r is not None,
            attempts=RETRY_ATTEMPTS,
            backoff_s=RETRY_BACKOFF_S,
            log=log,
            label=label,
        )
        if df is None:
            erros.append(label)
        elif df.empty:
            sem_dados.append(label)
        else:
            frames.append(df)
    return frames, sem_dados, erros


# ── Playwright helpers ──────────────────────────────────────────────────────
def first_visible(page, selector: str):
    """Devolve o primeiro elemento visível de um grupo de selectores, ou None."""
    loc = page.locator(selector)
    for i in range(loc.count()):
        el = loc.nth(i)
        try:
            if el.is_visible():
                return el
        except Exception:
            continue
    return None


# ── Moeda ───────────────────────────────────────────────────────────────────
# Prefixos de moeda → código ISO. Ordem importa (mais específicos primeiro).
CURRENCY_PREFIXES: list[tuple[str, str]] = [
    ("R$", "BRL"),
    ("€", "EUR"),
    ("£", "GBP"),
    ("COP", "COP"),
    ("MXN", "MXN"),
    ("PEN", "PEN"),
    ("ARS", "ARS"),
    ("CLP", "CLP"),
    ("USD", "USD"),
    ("$", "USD"),
]


def detect_currency(val) -> str | None:
    """Código ISO da moeda a partir do símbolo/prefixo de um valor, ou None."""
    if val is None:
        return None
    s = str(val).strip()
    for prefix, code in CURRENCY_PREFIXES:
        if prefix in s:
            return code
    return None
