"""
scripts/scp_cellxpert.py

Extração mensal dos ganhos de cada operador na plataforma Cellxpert.

A Cellxpert é uma plataforma de afiliados em SPA (Angular) servida sob
vários domínios white-label (evoaffiliates, blaze.partners, betobet.online,
bet90partners, …). O fluxo de login e de relatório é, porém, comum a todos.

A plataforma já é servida toda em inglês, por isso não é necessário (nem
desejável) forçar o idioma — fazê-lo via query/localStorage chegou a quebrar
a navegação para a página de login.

Fluxo por operador (linha activa do config/logins.xlsx):
1. Abre browser Playwright e faz login (selectores padrão Cellxpert)
2. Navega para o relatório  (— a definir nas próximas instruções —)
3. Extrai a tabela de ganhos
4. Padroniza num DataFrame e guarda um CSV por operador em data/
5. Concatena tudo num único ficheiro report/cellxpert.csv

Credenciais:
- URLs, usernames, operadores e flags vêm de  config/logins.xlsx
- Passwords vêm SEMPRE do .env  (var PASS_CELLXPERT_<OPERADOR>_<USERNAME>)
  Nunca são lidas da coluna Password do Excel.

Uso:
    python scripts/scp_cellxpert.py                    # todos os operadores activos
    python scripts/scp_cellxpert.py --operador BetsAmigo
    python scripts/scp_cellxpert.py --headful          # ver o browser a trabalhar
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
PLATFORM_SLUG = "CELLXPERT"
CURRENCY_PREFIXES = common.CURRENCY_PREFIXES

DEFAULT_TIMEOUT = common.DEFAULT_TIMEOUT
SLOW_MO = common.SLOW_MO
HEADLESS = common.HEADLESS

log = common.get_logger("scp_cellxpert")


# ── Config / contas ───────────────────────────────────────────────────────────
def env_key_for(operador: str, username: str) -> str:
    """Convenção da password no .env: PASS_CELLXPERT_<OPERADOR>_<USERNAME>."""
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
    """Lê os operadores Cellxpert activos do config/logins.xlsx."""
    df = common.load_login_rows("cellxpert", operador_filter, ids)
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


# ── Login Cellxpert ───────────────────────────────────────────────────────────
# A SPA da Cellxpert varia ligeiramente de tema, por isso usamos vários
# candidatos por campo e escolhemos o primeiro visível.
SEL_USERNAME = (
    "input[name='username'], input#username, input[formcontrolname='username'], "
    "input[type='text'][name*='user' i], input[name='email'], input[type='email']"
)
SEL_PASSWORD = (
    "input[name='password'], input#password, input[formcontrolname='password'], "
    "input[type='password']"
)
SEL_SUBMIT = (
    "button[type='submit'], input[type='submit'], button.login-button, "
    "button:has-text('Login'), button:has-text('Log in'), button:has-text('Sign in')"
)

# Selectores de erro (vários temas/idiomas).
_ERROR_SELECTORS = (
    ".alert-danger, .alert.alert-danger, .error, .error-message, .login-error, "
    ".validation-error, .mat-error, [class*='error']:not(input)"
)


def _first_visible(page: Page, selector: str):
    return common.first_visible(page, selector)


def login(page: Page, acc: Account) -> bool:
    """Login padrão Cellxpert. Devolve True em caso de sucesso."""
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
        form_visivel = _first_visible(page, SEL_USERNAME) is not None
        if not form_visivel:
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


def _login_error_text(page: Page) -> str:
    """Recolhe o texto dos alertas de erro visíveis na página de login."""
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


# ── Popup de boas-vindas ────────────────────────────────────────────────────────
# Ao entrar aparece um modal "Welcome" que bloqueia a navegação. Marcamos a opção
# "Don't show this message again" e clicamos em OK para não reaparecer nas próximas
# execuções da mesma conta.
SEL_WELCOME_CHECKBOX = "#do_not_show_again"
SEL_WELCOME_OK = "button:has-text('OK')"


def dismiss_welcome_popup(page: Page) -> None:
    """Fecha o popup 'Welcome', se presente, marcando 'não mostrar de novo'."""
    try:
        ok_btn = page.locator(SEL_WELCOME_OK).first
        ok_btn.wait_for(state="visible", timeout=5000)
    except Exception:
        return  # popup não apareceu

    try:
        checkbox = page.locator(SEL_WELCOME_CHECKBOX).first
        if checkbox.count() > 0:
            checkbox.check(force=True)  # input é estilizado/oculto atrás do span
    except Exception:
        pass

    try:
        page.locator(SEL_WELCOME_OK).first.click()
        page.wait_for_timeout(500)
        log.info("Popup de boas-vindas fechado")
    except Exception as e:
        log.debug(f"Não foi possível fechar o popup de boas-vindas: {e}")
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass


# ── Navegação para o Earnings Report ───────────────────────────────────────────
# O href é estável; as classes do menu são hashes voláteis de CSS-in-JS. Por isso
# navegamos direto pela URL, com fallback para o clique no menu se necessário.
EARNINGS_PATH = "/partner/reports/earnings"
SEL_REPORTS_SECTION = "section[data-tour='reports']"
SEL_EARNINGS_LINK = f"a[href='{EARNINGS_PATH}']"


def goto_earnings_report(page: Page) -> bool:
    """Abre o Earnings Report (URL directa; fallback via menu). True se navegou."""
    p = urlparse(page.url)
    earnings_url = f"{p.scheme}://{p.netloc}{EARNINGS_PATH}"
    try:
        page.goto(earnings_url, wait_until="networkidle", timeout=30000)
        if EARNINGS_PATH in page.url:
            log.info("Earnings Report aberto (URL directa)")
            return True
    except Exception as e:
        log.debug(f"Navegação directa falhou ({e}); a tentar via menu")

    # Fallback: expande a secção Reports e clica no link.
    try:
        section = page.locator(SEL_REPORTS_SECTION)
        section.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        section.locator("div", has_text="Reports").first.click()
        link = page.locator(SEL_EARNINGS_LINK)
        link.wait_for(state="visible", timeout=5000)
        link.click()
        page.wait_for_load_state("networkidle", timeout=30000)
        log.info("Earnings Report aberto (via menu)")
        return True
    except Exception as e:
        log.error(f"Falha ao abrir o Earnings Report: {e}")
        return False


# ── Selecção do período (mês anterior) ──────────────────────────────────────────
# O filtro de datas é um Ant Design RangePicker. Em vez de navegar célula a célula
# do calendário, abrimos o popup e usamos o preset "Previous Month".
SEL_PICKER = ".ant-picker"
SEL_PICKER_PANEL = ".ant-picker-panel-container"
SEL_PRESETS = ".ant-picker-presets"


def select_previous_month(page: Page) -> bool:
    """Abre o RangePicker e aplica o preset 'Previous Month'. True se aplicado."""
    try:
        picker = page.locator(SEL_PICKER).first
        picker.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        picker.click()

        # Espera o popup do calendário com a lista de presets.
        page.wait_for_selector(SEL_PICKER_PANEL, state="visible", timeout=DEFAULT_TIMEOUT)
        # Texto exacto para não confundir com 'Previous Week'/'Previous Quarter'.
        preset = page.locator(f"{SEL_PRESETS} li").get_by_text(
            "Previous Month", exact=True
        )
        preset.first.wait_for(state="visible", timeout=5000)
        preset.first.click()

        # O popup fecha após escolher o preset.
        page.wait_for_selector(SEL_PICKER_PANEL, state="hidden", timeout=10000)
        log.info("Período definido para 'Previous Month'")
        return True
    except Exception as e:
        log.error(f"Falha ao seleccionar o mês anterior: {e}")
        return False


def run_report(page: Page) -> bool:
    """Clica no botão 'Run Report' para executar o relatório com os filtros."""
    try:
        btn = page.get_by_role("button", name="Run Report")
        btn.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        btn.click()
        page.wait_for_load_state("networkidle", timeout=30000)
        log.info("Run Report executado")
        return True
    except Exception as e:
        log.error(f"Falha ao executar 'Run Report': {e}")
        return False


# ── Espera da tabela ────────────────────────────────────────────────────────────
SEL_NO_RESULTS = "text='No results found'"


class NoResults(Exception):
    """O relatório não devolveu dados para o período."""


def wait_for_table(page: Page) -> bool:
    """
    Aguarda, após o Run Report, ou a tabela de resultados ou a mensagem de
    'No results found'. Devolve True se há tabela; lança NoResults se vazio.
    """
    deadline = time.monotonic() + DEFAULT_TIMEOUT / 1000
    while time.monotonic() < deadline:
        if page.locator(SEL_NO_RESULTS).count() > 0:
            log.info("Sem resultados para o período — a passar à conta seguinte")
            raise NoResults
        # As linhas de dados têm a classe ant-table-row (a 1ª tr do tbody é uma
        # 'ant-table-measure-row' com height 0px / aria-hidden, nunca visível).
        rows = page.locator("table tbody tr.ant-table-row")
        if rows.count() > 0 and rows.first.is_visible():
            page.wait_for_load_state("networkidle", timeout=15000)
            log.info("Tabela de resultados carregada")
            return True
        page.wait_for_timeout(500)

    log.error("Timeout à espera da tabela/mensagem de resultados")
    return False


# ── Export CSV ──────────────────────────────────────────────────────────────────
def export_csv(page: Page) -> pd.DataFrame | None:
    """Clica em 'Export' → 'CSV' e devolve o DataFrame (lê o download do temp)."""
    try:
        # O Export é um dropdown radix (div[data-cy='export-button']), não um button.
        export_btn = page.locator("[data-cy='export-button']").first
        export_btn.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)
        export_btn.click()

        # O menu popup (radix dialog) abre com as opções; clicamos em CSV.
        csv_item = page.get_by_text("CSV", exact=True)
        csv_item.first.wait_for(state="visible", timeout=5000)
        with page.expect_download(timeout=30000) as dl_info:
            csv_item.first.click()
        # Lê o ficheiro temporário do download (não polui data/ com o bruto).
        df = pd.read_csv(dl_info.value.path(), dtype=str)
        log.info("CSV exportado")
        return df
    except Exception as e:
        log.error(f"Falha ao exportar CSV: {e}")
        return None


def _safe_name(name: str) -> str:
    return common.safe_name(name)


# ── Detecção de moeda ───────────────────────────────────────────────────────────
def _detect_currency(val) -> str | None:
    return common.detect_currency(val)


def capture_currency(page: Page) -> str | None:
    """Lê o símbolo da coluna 'Amount' na tabela (3ª coluna) e mapeia para ISO."""
    try:
        cells = page.locator("table tbody tr.ant-table-row td:nth-child(3)")
        for i in range(min(cells.count(), 10)):
            code = _detect_currency(cells.nth(i).inner_text())
            if code:
                return code
        # Fallback: linha de Total no rodapé (ex: €7,470.00)
        total = page.locator("tfoot.ant-table-summary td")
        for i in range(total.count()):
            code = _detect_currency(total.nth(i).inner_text())
            if code:
                return code
    except Exception as e:
        log.debug(f"Não foi possível detectar a moeda: {e}")
    return None


# ── Padronização e gravação ─────────────────────────────────────────────────────
def _previous_month_label() -> str:
    return common.previous_month_label()


# Tipos de comissão garantidos no output (mesmo sem dados, ficam com amount 0).
DEFAULT_COMMISSION_TYPES = common.DEFAULT_COMMISSION_TYPES


def _add_meta(grp: pd.DataFrame, acc: Account, currency: str | None) -> pd.DataFrame:
    """Acrescenta as colunas de identificação ao DataFrame agregado."""
    meta = {
        "operador": acc.operador,
        "empresa": acc.empresa,
        "plataforma": acc.plataforma,
        "username": acc.username,
        "month": _previous_month_label(),
        "currency": currency or "",
    }
    for i, (k, v) in enumerate(meta.items()):
        grp.insert(i, k, v)
    grp["extracted_at"] = datetime.now().isoformat(timespec="seconds")
    return grp


def empty_report(acc: Account, currency: str | None = None) -> pd.DataFrame:
    """Relatório para contas sem dados: CPA e Revshare Ongoing PL com amount 0."""
    grp = pd.DataFrame(
        {"commission_type": DEFAULT_COMMISSION_TYPES, "amount": [0.0, 0.0]}
    )
    return _add_meta(grp, acc, currency)


def standardize(df: pd.DataFrame, acc: Account, currency: str | None) -> pd.DataFrame:
    """Agrega o relatório por Commission Type e acrescenta colunas de identificação."""
    df = df.copy()
    df.columns = [str(c).strip().lstrip("﻿") for c in df.columns]
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)

    grp = (
        df.groupby("Commission Type", dropna=False)["Amount"]
        .sum()
        .reset_index()
        .rename(columns={"Commission Type": "commission_type", "Amount": "amount"})
    )
    grp["commission_type"] = grp["commission_type"].fillna("").astype(str).str.strip()
    grp["amount"] = grp["amount"].round(2)
    return _add_meta(grp, acc, currency)


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
        frames, REPORT_DIR / "cellxpert.csv", ["operador", "username", "month"], log, as_str=True
    )


# ── Extracção do relatório ─────────────────────────────────────────────────────
def extract_report(page: Page, acc: Account) -> pd.DataFrame | None:
    if not goto_earnings_report(page):
        return None
    # Algumas contas mostram o popup de boas-vindas também aqui.
    dismiss_welcome_popup(page)
    if not select_previous_month(page):
        return None
    if not run_report(page):
        return None
    try:
        if not wait_for_table(page):
            return None
    except NoResults:
        # Sem dados: gera linhas CPA=0 e Revshare Ongoing PL=0.
        return empty_report(acc)

    currency = capture_currency(page)
    raw = export_csv(page)
    if raw is None:
        return None
    if raw.empty:
        return empty_report(acc, currency)
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
            dismiss_welcome_popup(page)
            df = extract_report(page, acc)
            if df is None:
                return None  # falha real
            if df.empty:
                # Sem resultados para o período — não é erro; segue para a próxima.
                return df
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
        description="Extração mensal de ganhos dos operadores Cellxpert",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--operador", help="Processa apenas este operador (ex: BetsAmigo)")
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
        log.warning("Nenhum operador Cellxpert activo encontrado — a terminar")
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
