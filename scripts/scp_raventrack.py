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
import os
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
from playwright.sync_api import BrowserContext, Page, sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeout

try:  # funciona tanto em execução directa quanto importado como scripts.*
    import common
except ImportError:
    from scripts import common

common.load_env()

# ── Caminhos do projecto ────────────────────────────────────────────────────
ROOT = common.ROOT
DATA_DIR = common.DATA_DIR
REPORT_DIR = common.REPORT_DIR

# ── Constantes da plataforma ────────────────────────────────────────────────
PLATFORM_SLUG = "RAVENTRACK"
CURRENCY_PREFIXES = common.CURRENCY_PREFIXES

DEFAULT_TIMEOUT = common.DEFAULT_TIMEOUT
SLOW_MO = common.SLOW_MO
HEADLESS = common.HEADLESS

log = common.get_logger("scp_raventrack")


# ── Config / contas ───────────────────────────────────────────────────────────
def env_key_for(operador: str, username: str) -> str:
    """Convenção da password no .env: PASS_RAVENTRACK_<OPERADOR>_<USERNAME>."""
    return common.env_key_for(PLATFORM_SLUG, operador, username)


@dataclass
class Account:
    operador: str
    empresa: str
    plataforma: str
    username: str
    login_url: str
    file_name: str
    id: str = ""

    @property
    def password(self) -> str:
        return os.getenv(env_key_for(self.operador, self.username), "").strip()

    @property
    def base_url(self) -> str:
        return common.base_url_of(self.login_url)


def load_accounts(
    operador_filter: str | None = None, ids: list[str] | None = None
) -> list[Account]:
    """Lê os operadores RavenTrack activos do config/logins.xlsx."""
    df = common.load_login_rows("raventrack", operador_filter, ids)
    accounts: list[Account] = []
    for _, row in df.iterrows():
        url = str(row.get("URL", "")).strip()
        username = str(row.get("Username", "")).strip()
        operador = str(row.get("Operador", "")).strip()
        if not url or not username or url.lower() == "nan":
            continue
        accounts.append(
            Account(
                id=str(row.get("Id", "")).strip(),
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
    return common.first_visible(page, selector)


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
    except PlaywrightTimeout:
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
    return common.previous_month_label()


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
    return common.detect_currency(val)


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
    return common.safe_name(name)


def _upsert_csv(path: Path, new_df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    return common.upsert_csv(path, new_df, key_cols, log=log)


def save_operator_csv(df: pd.DataFrame, acc: Account) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / f"{_safe_name(acc.file_name)}.csv"
    merged = _upsert_csv(path, df.astype(str), ["operador", "username", "month"])
    merged.to_csv(path, index=False, encoding="utf-8-sig")
    log.info(f"Guardado {len(df)} linha(s) → {path.name}")
    return path


def save_combined(frames: list[pd.DataFrame]) -> Path | None:
    return common.save_combined(
        frames, REPORT_DIR / "raventrack.csv", ["operador", "username", "month"], log, as_str=True
    )


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
        "--id", nargs="+", dest="ids",
        help="Processa apenas a(s) conta(s) pela coluna Id do logins.xlsx",
    )
    parser.add_argument(
        "--headful", action="store_true", help="Mostra o browser (ignora HEADLESS=true)"
    )
    args = parser.parse_args()

    headless = HEADLESS and not args.headful

    log.info("=" * 60)
    log.info(f"Início — {datetime.now():%Y-%m-%d %H:%M:%S}")

    accounts = load_accounts(operador_filter=args.operador, ids=args.ids)
    if not accounts:
        log.warning("Nenhum operador RavenTrack activo encontrado — a terminar")
        sys.exit(0)
    log.info(f"{len(accounts)} operador(es) a processar")

    frames, sem_dados, erros = common.run_accounts(
        accounts, lambda acc: process_account(acc, headless=headless), log
    )

    save_combined(frames)

    log.info("=" * 60)
    if sem_dados:
        log.info(f"{len(sem_dados)} operador(es) sem resultados: {', '.join(sem_dados)}")
    if erros:
        log.warning(f"{len(erros)} operador(es) com erro: {', '.join(erros)}")
    log.info("Finalizado")


if __name__ == "__main__":
    main()
