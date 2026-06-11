"""
scripts/scp_netrefer.py

Extração mensal dos ganhos de cada operador na plataforma Netrefer.

Fluxo por operador (linha activa do config/logins.xlsx):
1. Abre browser Playwright e faz login (selectores padrão Netrefer)
2. Força o idioma da conta para EN
3. Navega para o relatório  /affiliates/Earnings/MonthlyEarnings
4. Extrai a tabela de ganhos (download CSV ou leitura do HTML)
5. Padroniza num DataFrame e guarda um CSV por operador em data/
6. Concatena tudo num único ficheiro report/netrefer.csv

Credenciais:
- URLs, usernames, operadores e flags vêm de  config/logins.xlsx
- Passwords vêm SEMPRE do .env  (var PASS_NETREFER_<OPERADOR>_<USERNAME>)
  Nunca são lidas da coluna Password do Excel.

Uso:
    python scripts/scp_netrefer.py                 # todos os operadores activos
    python scripts/scp_netrefer.py --operador SNAI # só um operador
    python scripts/scp_netrefer.py --headful       # ver o browser a trabalhar
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
from io import StringIO
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
PLATFORM_SLUG = "NETREFER"
REPORT_PATH = "/affiliates/Earnings/MonthlyEarnings"
LANG_EN_PATH = "/affiliates/Home/UpdateUserLanguage?languageID=1"  # 1 = English

# Prefixos de moeda → código ISO. Ordem importa (mais longos/específicos primeiro).
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
def get_logger(name: str = "scp_netrefer") -> logging.Logger:
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
    """Convenção da password no .env: PASS_NETREFER_<OPERADOR>_<USERNAME>."""
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
    """Lê os operadores activos do config/logins.xlsx."""
    if not CONFIG_XLSX.exists():
        log.error(f"Ficheiro de config não encontrado: {CONFIG_XLSX}")
        return []

    df = pd.read_excel(CONFIG_XLSX, dtype=str)
    df = df[
        (df["Plataforma"].str.strip().str.lower() == "netrefer")
        & (df["Active"].astype(str).str.strip() == "1")
    ].reset_index(drop=True)

    accounts: list[Account] = []
    for _, row in df.iterrows():
        url = str(row.get("URL", "")).strip().rstrip("#").strip()
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
                file_name=str(row.get("FileName", "")).strip() or f"{operador}_{username}",
            )
        )
    return accounts


# ── Login Netrefer ──────────────────────────────────────────────────────────
SEL_USERNAME = "#txtUsername"
SEL_PASSWORD = "#txtPassword"
SEL_SUBMIT = "#btnLogin"
SEL_AGREE = "#agreeButton"


def login(page: Page, acc: Account) -> bool:
    """Login padrão Netrefer. Devolve True em caso de sucesso."""
    log.info(f"A fazer login em {acc.base_url}")
    page.goto(acc.login_url, wait_until="networkidle", timeout=30000)

    page.wait_for_selector(SEL_USERNAME, timeout=DEFAULT_TIMEOUT)
    page.fill(SEL_USERNAME, acc.username)
    page.fill(SEL_PASSWORD, acc.password)
    page.click(SEL_SUBMIT)
    page.wait_for_load_state("domcontentloaded", timeout=30000)

    _handle_agree_popup(page)

    # Detecção robusta: faz poll por sucesso (links autenticados / saída do form)
    # vs. erro explícito, em vez de um wait fixo curto que gera falsos negativos
    # em logins lentos.
    auth_selector = "a[href*='Earnings'], a[href*='Reports'], a[href*='logout']"
    deadline = time.monotonic() + 30  # até 30s para concluir o login
    while time.monotonic() < deadline:
        _handle_agree_popup(page)

        if page.locator(auth_selector).count() > 0:
            log.info("Login OK")
            return True

        form_visivel = (
            page.locator(SEL_USERNAME).count() > 0
            and page.locator(SEL_USERNAME).is_visible()
        )
        if form_visivel:
            erro = _login_error_text(page)
            if erro:  # erro explícito do site → falha definitiva, não espera mais
                log.error(f"Login falhou para {acc.operador}/{acc.username}: {erro}")
                return False
            # form ainda visível mas sem erro → provavelmente a carregar; aguarda
        elif "login" not in page.url.lower():
            # Saímos da página de login e o form desapareceu → sucesso
            log.info("Login OK")
            return True

        page.wait_for_timeout(1000)

    motivo = _login_error_text(page) or "timeout (30s) sem confirmar sessão autenticada"
    log.error(f"Login falhou para {acc.operador}/{acc.username}: {motivo}")
    return False


# Selectores de alertas de erro do Netrefer (vários temas/idiomas)
_ERROR_SELECTORS = (
    ".alert-danger, .alert.alert-danger, #login-error, .login-error, "
    ".error-message, .validation-summary-errors, .alert-error, "
    ".form-error, #errorMessage, .field-validation-error"
)


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
        return " | ".join(textos).strip()
    except Exception:
        return ""


def _handle_agree_popup(page: Page) -> None:
    """Aceita popup de sessão activa e/ou página de termos e condições."""
    try:
        page.wait_for_selector(SEL_AGREE, timeout=3000)
        page.click(SEL_AGREE)
        page.wait_for_load_state("domcontentloaded", timeout=DEFAULT_TIMEOUT)
    except Exception:
        pass
    try:
        tos = page.locator(
            "input[value*='I agree to the terms'], input.btn-blue[type='submit']"
        )
        if tos.count() > 0 and tos.first.is_visible():
            log.info("Página de termos detectada — a aceitar...")
            tos.first.click()
            page.wait_for_load_state("domcontentloaded", timeout=DEFAULT_TIMEOUT)
    except Exception:
        pass


def force_english(page: Page) -> None:
    """Força a conta para EN via fetch silencioso (sem carregar página extra)."""
    try:
        page.evaluate(f"fetch('{LANG_EN_PATH}')")
        log.debug("Idioma definido para EN")
    except Exception as e:
        log.debug(f"Não foi possível forçar EN: {e}")


# ── Extracção do relatório Monthly Earnings ───────────────────────────────────
def extract_monthly_earnings(page: Page, acc: Account) -> pd.DataFrame | None:
    """
    Navega para  /affiliates/Earnings/MonthlyEarnings  e devolve a tabela
    de ganhos como DataFrame. Não é necessário seleccionar período.

    Estratégia de extracção (na ordem):
      1. Botão de download CSV do DataTables, se existir
      2. Leitura directa da(s) tabela(s) HTML renderizada(s)
    """
    report_url = acc.base_url + REPORT_PATH
    page.goto(report_url, wait_until="domcontentloaded")
    log.info(f"A navegar para: {report_url}")

    # Sessão expirou → voltou ao login
    if page.locator(SEL_USERNAME).count() > 0 and page.locator(SEL_USERNAME).is_visible():
        log.warning("Sessão expirou ao abrir o relatório — a tentar novo login...")
        if not login(page, acc):
            return None
        page.goto(report_url, wait_until="domcontentloaded")

    # Espera que uma tabela com linhas renderize
    try:
        page.wait_for_selector("table tbody tr", state="visible", timeout=DEFAULT_TIMEOUT)
    except Exception:
        log.warning("Timeout à espera da tabela do relatório")

    df = _download_csv(page, acc)
    if df is None:
        df = _scrape_html_table(page)
    if df is None or df.empty:
        log.warning(f"Sem dados extraídos para {acc.operador}/{acc.username}")
        return None
    return df


def _download_csv(page: Page, acc: Account) -> pd.DataFrame | None:
    """Tenta o botão CSV do DataTables. Guarda o ficheiro bruto e devolve DataFrame."""
    try:
        btn = page.locator("a.dt-button.buttons-csv, .dt-button.buttons-csv").first
        if btn.count() == 0:
            return None
        btn.wait_for(state="visible", timeout=4000)
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        raw_path = DATA_DIR / f"{_safe_name(acc.file_name)}.csv"
        with page.expect_download() as dl_info:
            btn.click()
        dl_info.value.save_as(raw_path)
        log.info(f"CSV bruto guardado: {raw_path.name}")
        return pd.read_csv(raw_path, dtype=str)
    except Exception as e:
        log.debug(f"Download CSV indisponível ({e}); a usar leitura de HTML")
        return None


def _scrape_html_table(page: Page) -> pd.DataFrame | None:
    """Lê a maior tabela HTML da página como fallback."""
    try:
        html = page.content()
        tables = pd.read_html(StringIO(html))
        if not tables:
            return None
        df = max(tables, key=len)
        df.columns = [str(c).strip() for c in df.columns]
        return df.astype(str)
    except Exception as e:
        log.debug(f"Falha a ler tabela HTML: {e}")
        return None


# ── Padronização e gravação ───────────────────────────────────────────────────
def standardize(df: pd.DataFrame, acc: Account) -> pd.DataFrame:
    """Limpa a tabela e acrescenta as colunas de identificação do operador."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Remove linha de totais, se existir (primeira coluna começa por "total")
    if len(df.columns) > 0:
        first = df.columns[0]
        df = df[
            ~df[first].astype(str).str.strip().str.lower().str.startswith("total")
        ].reset_index(drop=True)

    # Detecta a moeda pelos símbolos das colunas de valores (€, R$, $, £...).
    # Colunas de valores = todas excepto "Month".
    value_cols = [c for c in df.columns if c.strip().lower() != "month"]
    currency = df.apply(lambda r: detect_row_currency(r, value_cols), axis=1)

    from datetime import date as _date
    _today = _date.today().replace(day=1)
    _prev = (_today - __import__("datetime").timedelta(days=1)).replace(day=1)
    month_label = _prev.strftime("%Y-%m")

    meta = {
        "operador": acc.operador,
        "empresa": acc.empresa,
        "plataforma": acc.plataforma,
        "username": acc.username,
        "month": month_label,
        "currency": currency,
        "extracted_at": datetime.now().isoformat(timespec="seconds"),
    }
    for i, (k, v) in enumerate(meta.items()):
        df.insert(i, k, v)
    return df


def _detect_currency(val) -> str | None:
    """Devolve o código ISO da moeda a partir do símbolo/prefixo de um valor."""
    if val is None:
        return None
    s = str(val).strip().lstrip()
    for prefix, code in CURRENCY_PREFIXES:
        if prefix in s:
            return code
    return None


def detect_row_currency(row: pd.Series, value_cols: list[str]) -> str | None:
    """Detecta a moeda de uma linha varrendo as colunas de valores monetários."""
    for col in value_cols:
        code = _detect_currency(row.get(col))
        if code:
            return code
    return None


def _safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]", "_", name).strip("_") or "report"


def _upsert_csv(path: Path, new_df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    """Substitui no CSV existente as linhas cujas key_cols coincidem com new_df; appenda o resto."""
    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
            missing = [c for c in key_cols if c not in existing.columns]
            if missing:
                log.warning(f"Upsert: colunas chave ausentes no ficheiro existente: {missing} — a sobrescrever")
                return new_df.copy()
            mask = existing[key_cols].apply(tuple, axis=1).isin(
                new_df[key_cols].apply(tuple, axis=1)
            )
            removed = mask.sum()
            existing = existing[~mask]
            if removed:
                log.debug(f"Upsert: {removed} linha(s) substituída(s) em {path.name}")
            return pd.concat([existing, new_df], ignore_index=True)
        except Exception as e:
            log.warning(f"Upsert falhou ({e}) — a sobrescrever {path.name}")
    return new_df.copy()


def save_operator_csv(df: pd.DataFrame, acc: Account) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = DATA_DIR / f"{_safe_name(acc.file_name)}.csv"
    merged = _upsert_csv(path, df, ["month"])
    merged.to_csv(path, index=False, encoding="utf-8-sig")
    log.info(f"Guardado {len(df)} linhas → {path.name}")
    return path


def save_combined(frames: list[pd.DataFrame]) -> Path | None:
    if not frames:
        log.warning("Nenhum dado para consolidar — report/netrefer.csv não gerado")
        return None
    new_data = pd.concat(frames, ignore_index=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    path = REPORT_DIR / "netrefer.csv"
    merged = _upsert_csv(path, new_data, ["operador", "username", "month"])
    merged.to_csv(path, index=False, encoding="utf-8-sig")
    log.info(f"Ficheiro consolidado: {path}  ({len(merged)} linhas)")
    return path


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
        context: BrowserContext = browser.new_context(accept_downloads=True)
        page: Page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)
        try:
            if not login(page, acc):
                return None
            force_english(page)
            df = extract_monthly_earnings(page, acc)
            if df is None:
                return None
            df = standardize(df, acc)
            save_operator_csv(df, acc)
            return df
        except Exception as e:
            log.exception(f"Erro inesperado em {acc.operador}/{acc.username}: {e}")
            return None
        finally:
            browser.close()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extração mensal de ganhos dos operadores Netrefer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--operador", help="Processa apenas este operador (ex: SNAI)")
    parser.add_argument(
        "--headful", action="store_true", help="Mostra o browser (ignora HEADLESS=true)"
    )
    args = parser.parse_args()

    headless = HEADLESS and not args.headful

    log.info("=" * 60)
    log.info(f"Início — {datetime.now():%Y-%m-%d %H:%M:%S}")

    accounts = load_accounts(operador_filter=args.operador)
    if not accounts:
        log.warning("Nenhum operador activo encontrado — a terminar")
        sys.exit(0)
    log.info(f"{len(accounts)} operador(es) a processar")

    frames: list[pd.DataFrame] = []
    erros: list[str] = []
    for i, acc in enumerate(accounts, 1):
        log.info(f"[{i}/{len(accounts)}] {acc.operador} / {acc.username}")
        df = process_account(acc, headless=headless)
        if df is not None and not df.empty:
            frames.append(df)
        else:
            erros.append(f"{acc.operador}/{acc.username}")

    save_combined(frames)

    log.info("=" * 60)
    if erros:
        log.warning(f"{len(erros)} operador(es) sem dados/erro: {', '.join(erros)}")
    log.info("Finalizado")


if __name__ == "__main__":
    main()
