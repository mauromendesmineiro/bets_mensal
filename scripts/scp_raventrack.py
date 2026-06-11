"""
scripts/scp_raventrack.py

Extração mensal dos ganhos de cada operador na plataforma RavenTrack
(plataforma Raven5 / "Affiliates", servida sob vários domínios white-label:
partner.novibet.com, webaffiliates-rt.yajuego.co, …).

Fluxo por operador (linha activa do config/logins.xlsx):
1. Abre browser Playwright e faz login            (raventrack_login.txt)
2. Abre o relatório de Vendor                     (raventrack_report.txt / raventrack_vendor.txt)
3. Selecciona o período (mês anterior)            (raventrack_date_range.txt)
4. Clica em Search                                (raventrack_search.txt)
5. Lê a tabela de resultados                      (raventrack_table.txt)
6. Clica em "Export All" e descarrega o ficheiro  (raventrack_export.txt)
7. Padroniza num DataFrame e guarda um CSV por operador em data/
8. Concatena tudo num único ficheiro report/raventrack.csv

Nota sobre a moeda: o ficheiro do "Export All" costuma vir sem o símbolo da
moeda (como no Cellxpert), por isso os dados vêm do export (que garante todas
as linhas, mesmo paginadas) mas a moeda é lida da tabela renderizada
(#results-table), que mostra os valores COM o símbolo (ex.: €21.47).

Credenciais:
- URLs, usernames, operadores e flags vêm de  config/logins.xlsx
- Passwords vêm SEMPRE do .env  (var PASS_RAVENTRACK_<OPERADOR>_<USERNAME>)
  Nunca são lidas da coluna Password do Excel.

Uso:
    python scripts/scp_raventrack.py                  # todos os operadores activos
    python scripts/scp_raventrack.py --operador Novibet
    python scripts/scp_raventrack.py --headful        # ver o browser a trabalhar
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
from playwright.sync_api import BrowserContext, Page, sync_playwright

# ── Caminhos do projecto ────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parents[1]
CONFIG_XLSX = ROOT / "config" / "logins.xlsx"
DATA_DIR = ROOT / "data"
REPORT_DIR = ROOT / "report"
LOGS_DIR = ROOT / "logs"


def _load_env() -> None:
    """Carrega .env suportando chaves entre aspas ("KEY"="VALUE")."""
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


_load_env()

# ── Constantes da plataforma ────────────────────────────────────────────────
PLATFORM_SLUG = "RAVENTRACK"

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

DEFAULT_TIMEOUT = int(os.getenv("DEFAULT_TIMEOUT", "20000"))
SLOW_MO = int(os.getenv("SLOW_MO", "0"))
HEADLESS = os.getenv("HEADLESS", "true").lower() == "true"


# ── Logger ──────────────────────────────────────────────────────────────────
def get_logger(name: str = "scp_raventrack") -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)-7s] %(message)s", datefmt="%H:%M:%S"
    )
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
        fh = logging.FileHandler(
            LOGS_DIR / f"{datetime.now():%Y-%m-%d}_{plat}.log", encoding="utf-8"
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        pass
    return logger


log = get_logger()


# ── Config / contas ───────────────────────────────────────────────────────────
def sanitize(val: str) -> str:
    """Normaliza um valor para compor nomes de variáveis de ambiente."""
    return re.sub(r"[^A-Z0-9]", "_", str(val).upper())


def env_key_for(operador: str, username: str) -> str:
    """Convenção da password no .env: PASS_RAVENTRACK_<OPERADOR>_<USERNAME>."""
    return f"PASS_{PLATFORM_SLUG}_{sanitize(operador)}_{sanitize(username)}"


@dataclass
class Account:
    operador: str
    empresa: str
    plataforma: str
    username: str
    login_url: str
    file_name: str

    @property
    def password(self) -> str:
        return os.getenv(env_key_for(self.operador, self.username), "").strip()

    @property
    def base_url(self) -> str:
        p = urlparse(self.login_url)
        return f"{p.scheme}://{p.netloc}"


def load_accounts(operador_filter: str | None = None) -> list[Account]:
    """Lê os operadores RavenTrack activos do config/logins.xlsx."""
    if not CONFIG_XLSX.exists():
        log.error(f"Ficheiro de config não encontrado: {CONFIG_XLSX}")
        return []

    df = pd.read_excel(CONFIG_XLSX, dtype=str)
    df = df[
        (df["Plataforma"].str.strip().str.lower() == "raventrack")
        & (df["Active"].astype(str).str.strip() == "1")
    ].reset_index(drop=True)

    accounts: list[Account] = []
    for _, row in df.iterrows():
        url = str(row.get("URL", "")).strip()
        username = str(row.get("Username", "")).strip()
        operador = str(row.get("Operador", "")).strip()
        if not url or not username or url.lower() == "nan":
            continue
        if operador_filter and operador.lower() != operador_filter.lower():
            continue
        accounts.append(
            Account(
                operador=operador,
                empresa=str(row.get("Empresa", "")).strip(),
                plataforma=str(row.get("Plataforma", "")).strip(),
                username=username,
                login_url=url,
                file_name=str(row.get("FileName", "")).strip()
                or f"{operador}_{username}",
            )
        )
    return accounts


# ── Login RavenTrack ──────────────────────────────────────────────────────────
# O username é um e-mail. O tema varia por domínio white-label, por isso usamos
# vários candidatos por campo e escolhemos o primeiro visível.
SEL_USERNAME = (
    "input[name='email'], input#email, input[type='email'], "
    "input[name='username'], input#username, input[type='text'][name*='user' i]"
)
SEL_PASSWORD = (
    "input[name='password'], input#password, input[type='password']"
)
SEL_SUBMIT = (
    "button[type='submit'], input[type='submit'], button.card-login, "
    "button:has-text('Login'), button:has-text('Log in'), button:has-text('Sign in')"
)
_ERROR_SELECTORS = (
    ".alert-danger, .alert.alert-danger, .invalid-feedback, .error, "
    ".error-message, .help-block, [class*='error']:not(input)"
)


def _first_visible(page: Page, selector: str):
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


def _login_error_text(page: Page) -> str:
    try:
        els = page.locator(_ERROR_SELECTORS)
        textos = []
        for i in range(els.count()):
            el = els.nth(i)
            try:
                if el.is_visible():
                    t = (el.text_content() or "").strip()
                    if t:
                        textos.append(t)
            except Exception:
                pass
        return " | ".join(dict.fromkeys(textos)).strip()
    except Exception:
        return ""


def login(page: Page, acc: Account) -> bool:
    """Login padrão Raven5. Devolve True em caso de sucesso."""
    log.info(f"A fazer login em {acc.base_url}")
    page.goto(acc.login_url, wait_until="domcontentloaded", timeout=30000)

    try:
        page.wait_for_selector(SEL_USERNAME, state="visible", timeout=DEFAULT_TIMEOUT)
    except Exception:
        log.error(f"Campo de username não apareceu em {acc.base_url}")
        return False

    user_el = _first_visible(page, SEL_USERNAME)
    pass_el = _first_visible(page, SEL_PASSWORD)
    if not user_el or not pass_el:
        log.error("Não foi possível localizar os campos de login")
        return False

    user_el.fill(acc.username)
    pass_el.fill(acc.password)

    submit_el = _first_visible(page, SEL_SUBMIT)
    if submit_el:
        submit_el.click()
    else:
        pass_el.press("Enter")
    page.wait_for_load_state("domcontentloaded", timeout=30000)

    # Poll por sucesso (form de login desapareceu) vs. erro explícito.
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        if _first_visible(page, SEL_USERNAME) is None and "/account/login" not in page.url:
            log.info("Login OK")
            return True
        erro = _login_error_text(page)
        if erro:
            log.error(f"Login falhou para {acc.operador}/{acc.username}: {erro}")
            return False
        page.wait_for_timeout(1000)

    motivo = _login_error_text(page) or "timeout (30s) sem confirmar sessão autenticada"
    log.error(f"Login falhou para {acc.operador}/{acc.username}: {motivo}")
    return False


# ── Navegação para o relatório de Vendor ────────────────────────────────────────
# O caminho é estável; navegamos directo pela URL com fallback via menu.
VENDOR_PATH = "/reporting/vendor"
SEL_REPORTING_LINK = "a[href*='/reporting/']:has-text('Reporting')"
SEL_VENDOR_LINK = f"a[href$='{VENDOR_PATH}']"


def goto_vendor_report(page: Page) -> bool:
    """Abre o relatório de Vendor (URL directa; fallback via menu)."""
    vendor_url = f"{_base(page)}{VENDOR_PATH}"
    try:
        page.goto(vendor_url, wait_until="networkidle", timeout=30000)
        if VENDOR_PATH in page.url:
            log.info("Relatório de Vendor aberto (URL directa)")
            return True
    except Exception as e:
        log.debug(f"Navegação directa falhou ({e}); a tentar via menu")

    # Fallback: clica em Reporting e depois no separador Vendor.
    try:
        link = page.locator(SEL_REPORTING_LINK).first
        link.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        link.click()
        page.wait_for_load_state("networkidle", timeout=30000)
        vendor = page.locator(SEL_VENDOR_LINK).first
        vendor.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        vendor.click()
        page.wait_for_load_state("networkidle", timeout=30000)
        log.info("Relatório de Vendor aberto (via menu)")
        return True
    except Exception as e:
        log.error(f"Falha ao abrir o relatório de Vendor: {e}")
        return False


def _base(page: Page) -> str:
    p = urlparse(page.url)
    return f"{p.scheme}://{p.netloc}"


# ── Selecção do período (mês anterior) ──────────────────────────────────────────
# É um <select id="date_range"> cujas opções de mês têm value no formato YYYY-MM
# (ex.: value="2026-05" → "May 2026").
SEL_DATE_RANGE = "select#date_range, select[name='date_range']"


def _previous_month_label() -> str:
    from datetime import date as _date, timedelta

    first = _date.today().replace(day=1)
    prev = (first - timedelta(days=1)).replace(day=1)
    return prev.strftime("%Y-%m")


def select_previous_month(page: Page) -> bool:
    """Selecciona o mês anterior no <select> de período (por value YYYY-MM)."""
    target = _previous_month_label()
    try:
        sel = page.locator(SEL_DATE_RANGE).first
        sel.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        sel.select_option(value=target)
        log.info(f"Período definido para {target}")
        return True
    except Exception as e:
        log.error(f"Falha ao seleccionar o mês anterior ({target}): {e}")
        return False


# ── Search ──────────────────────────────────────────────────────────────────────
SEL_SEARCH = "#analytics-search-button"


def run_search(page: Page) -> bool:
    """Clica em 'Search' para correr o relatório com os filtros escolhidos."""
    try:
        btn = page.locator(SEL_SEARCH).first
        btn.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        btn.click()
        page.wait_for_load_state("networkidle", timeout=30000)
        log.info("Search executado")
        return True
    except Exception as e:
        log.error(f"Falha ao executar 'Search': {e}")
        return False


# ── Leitura da tabela de resultados ─────────────────────────────────────────────
SEL_RESULTS_TABLE = "#results-table"
SEL_RESULTS_ROWS = "#results-table tbody tr"


class NoResults(Exception):
    """O relatório não devolveu linhas para o período."""


def wait_for_table(page: Page) -> bool:
    """Aguarda a tabela de resultados ficar com pelo menos uma linha."""
    deadline = time.monotonic() + DEFAULT_TIMEOUT / 1000
    while time.monotonic() < deadline:
        try:
            page.wait_for_selector(SEL_RESULTS_TABLE, state="visible", timeout=1000)
        except Exception:
            continue
        rows = page.locator(SEL_RESULTS_ROWS)
        if rows.count() > 0 and rows.first.is_visible():
            log.info(f"Tabela de resultados carregada ({rows.count()} linha(s))")
            return True
        page.wait_for_timeout(500)
    log.warning("Tabela de resultados sem linhas dentro do tempo limite")
    raise NoResults


def scrape_table(page: Page) -> pd.DataFrame:
    """Lê o #results-table (cabeçalhos + linhas) para um DataFrame de texto."""
    headers = [
        (h.inner_text() or "").strip()
        for h in page.locator(f"{SEL_RESULTS_TABLE} thead th").all()
    ]
    rows_data: list[list[str]] = []
    for tr in page.locator(SEL_RESULTS_ROWS).all():
        cells = [(td.inner_text() or "").strip() for td in tr.locator("td").all()]
        if cells:
            rows_data.append(cells)

    if not rows_data:
        raise NoResults

    # Garante alinhamento entre nº de cabeçalhos e nº de colunas.
    width = max(len(headers), max(len(r) for r in rows_data))
    if len(headers) < width:
        headers += [f"col_{i}" for i in range(len(headers), width)]
    rows_data = [r + [""] * (width - len(r)) for r in rows_data]

    return pd.DataFrame(rows_data, columns=headers[:width])


# ── Export All (download) ───────────────────────────────────────────────────────
# Botão "Export All" (button.report-export). Garante todas as linhas, mesmo que a
# tabela on-page esteja paginada. O ficheiro exportado costuma vir sem o símbolo da
# moeda, por isso esta é apenas a fonte de dados — a moeda é lida da tabela.
SEL_EXPORT = "button.report-export, button:has-text('Export All')"


def export_all(page: Page) -> pd.DataFrame | None:
    """Clica em 'Export All' e devolve o DataFrame do ficheiro descarregado."""
    try:
        btn = page.locator(SEL_EXPORT).first
        btn.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        with page.expect_download(timeout=30000) as dl_info:
            btn.click()
        dl = dl_info.value
        path = dl.path()
        name = (dl.suggested_filename or "").lower()
        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(path, dtype=str)
        else:
            df = pd.read_csv(path, dtype=str)
        log.info(f"Export All concluído ({len(df)} linha(s))")
        return df
    except Exception as e:
        log.warning(f"Export All falhou ({e}) — fallback para leitura da tabela")
        return None


# ── Detecção de moeda ───────────────────────────────────────────────────────────
def _detect_currency(val) -> str | None:
    if val is None:
        return None
    s = str(val).strip()
    for prefix, code in CURRENCY_PREFIXES:
        if prefix in s:
            return code
    return None


def detect_currency_from_df(df: pd.DataFrame) -> str | None:
    """Procura o primeiro símbolo de moeda em qualquer célula do DataFrame."""
    for col in df.columns:
        for val in df[col].astype(str):
            code = _detect_currency(val)
            if code:
                return code
    return None


# ── Padronização e gravação ─────────────────────────────────────────────────────
def _add_meta(df: pd.DataFrame, acc: Account, currency: str | None) -> pd.DataFrame:
    """Acrescenta as colunas de identificação ao DataFrame da tabela."""
    df = df.copy()
    meta = {
        "operador": acc.operador,
        "empresa": acc.empresa,
        "plataforma": acc.plataforma,
        "username": acc.username,
        "month": _previous_month_label(),
        "currency": currency or "",
    }
    for i, (k, v) in enumerate(meta.items()):
        df.insert(i, k, v)
    df["extracted_at"] = datetime.now().isoformat(timespec="seconds")
    return df


def standardize(df: pd.DataFrame, acc: Account, currency: str | None) -> pd.DataFrame:
    """Limpa cabeçalhos e acrescenta as colunas de identificação."""
    df = df.copy()
    df.columns = [str(c).strip().lstrip("﻿") for c in df.columns]
    return _add_meta(df, acc, currency)


def _safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]", "_", name).strip("_") or "report"


def _upsert_csv(path: Path, new_df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    """Substitui no CSV existente as linhas cujas key_cols coincidem; appenda o resto."""
    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
            missing = [c for c in key_cols if c not in existing.columns]
            if missing:
                log.warning(
                    f"Upsert: colunas chave ausentes em {path.name}: {missing} — a sobrescrever"
                )
                return new_df.copy()
            mask = existing[key_cols].apply(tuple, axis=1).isin(
                new_df[key_cols].astype(str).apply(tuple, axis=1)
            )
            return pd.concat([existing[~mask], new_df], ignore_index=True)
        except Exception as e:
            log.warning(f"Upsert falhou ({e}) — a sobrescrever {path.name}")
    return new_df.copy()


def save_operator_csv(df: pd.DataFrame, acc: Account) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / f"{_safe_name(acc.file_name)}.csv"
    merged = _upsert_csv(path, df.astype(str), ["month"])
    merged.to_csv(path, index=False, encoding="utf-8-sig")
    log.info(f"Guardado {len(df)} linha(s) → {path.name}")
    return path


def save_combined(frames: list[pd.DataFrame]) -> Path | None:
    if not frames:
        log.warning("Nenhum dado para consolidar — report/raventrack.csv não gerado")
        return None
    new_data = pd.concat(frames, ignore_index=True).astype(str)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    path = REPORT_DIR / "raventrack.csv"
    # Substitui todo o período do par operador+username (regra de merge acordada).
    merged = _upsert_csv(path, new_data, ["operador", "username", "month"])
    merged.to_csv(path, index=False, encoding="utf-8-sig")
    log.info(f"Ficheiro consolidado: {path}  ({len(merged)} linhas)")
    return path


# ── Extracção do relatório ─────────────────────────────────────────────────────
def extract_report(page: Page, acc: Account) -> pd.DataFrame | None:
    if not goto_vendor_report(page):
        return None
    if not select_previous_month(page):
        return None
    if not run_search(page):
        return None
    try:
        wait_for_table(page)
        table_df = scrape_table(page)
    except NoResults:
        log.info("Sem resultados para o período — a passar à conta seguinte")
        return pd.DataFrame()  # sem dados (não é erro)

    # A moeda vem da tabela on-page (que tem o símbolo); os dados, do Export All.
    currency = detect_currency_from_df(table_df)
    raw = export_all(page)
    if raw is None or raw.empty:
        raw = table_df  # fallback: usa a própria tabela raspada
    if currency is None:
        currency = detect_currency_from_df(raw)
    return standardize(raw, acc, currency)


# ── Orquestrador ──────────────────────────────────────────────────────────────
def process_account(acc: Account, headless: bool) -> pd.DataFrame | None:
    if not acc.password:
        log.error(
            f"Password não encontrada no .env para {acc.operador}/{acc.username} "
            f"— var esperada: {env_key_for(acc.operador, acc.username)}"
        )
        return None

    log.info(f"── {acc.plataforma} / {acc.operador} / {acc.username}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=SLOW_MO)
        context: BrowserContext = browser.new_context(
            accept_downloads=True, locale="en-US"
        )
        page: Page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)
        try:
            if not login(page, acc):
                return None
            df = extract_report(page, acc)
            if df is None:
                return None  # falha real
            if df.empty:
                return df  # sem resultados — segue para a próxima
            save_operator_csv(df, acc)
            log.info(f"{len(df)} linha(s) extraída(s) para {acc.operador}/{acc.username}")
            return df
        except Exception as e:
            log.exception(f"Erro inesperado em {acc.operador}/{acc.username}: {e}")
            return None
        finally:
            browser.close()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extração mensal de ganhos dos operadores RavenTrack",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--operador", help="Processa apenas este operador (ex: Novibet)")
    parser.add_argument(
        "--headful", action="store_true", help="Mostra o browser (ignora HEADLESS=true)"
    )
    args = parser.parse_args()

    headless = HEADLESS and not args.headful

    log.info("=" * 60)
    log.info(f"Início — {datetime.now():%Y-%m-%d %H:%M:%S}")

    accounts = load_accounts(operador_filter=args.operador)
    if not accounts:
        log.warning("Nenhum operador RavenTrack activo encontrado — a terminar")
        sys.exit(0)
    log.info(f"{len(accounts)} operador(es) a processar")

    erros: list[str] = []
    sem_dados: list[str] = []
    frames: list[pd.DataFrame] = []
    for i, acc in enumerate(accounts, 1):
        log.info(f"[{i}/{len(accounts)}] {acc.operador} / {acc.username}")
        df = process_account(acc, headless=headless)
        if df is None:
            erros.append(f"{acc.operador}/{acc.username}")
        elif df.empty:
            sem_dados.append(f"{acc.operador}/{acc.username}")
        else:
            frames.append(df)

    save_combined(frames)

    log.info("=" * 60)
    if sem_dados:
        log.info(f"{len(sem_dados)} operador(es) sem resultados: {', '.join(sem_dados)}")
    if erros:
        log.warning(f"{len(erros)} operador(es) com erro: {', '.join(erros)}")
    log.info("Finalizado")


if __name__ == "__main__":
    main()
