"""
scripts/scp_income_access.py

Extração mensal dos ganhos de cada operador na plataforma Income Access.

Fluxo por operador (linha activa do config/logins.xlsx, Plataforma=IncomeAccess):
1. Abre browser Playwright e faz login (com ou sem captcha via 2captcha)
2. Navega para o relatório de comissões (Commission Report)
3. Configura Merchant = "All Merchants" e extrai o CSV
4. Padroniza num DataFrame e guarda um CSV por operador em data/
5. Concatena tudo num único ficheiro report/income_access.csv

Flags Active no logins.xlsx:
  1 → activa, sem captcha
  4 → activa, com captcha (Betano domínio principal)
  5 → activa, com captcha (Betano AR)
  0 → inactiva → ignorada
  3 → ignorada (Pokerstar — estrutura distinta, não suportada)

Credenciais:
  - URLs, usernames, operadores e flags vêm de config/logins.xlsx
  - Passwords → .env  (var PASS_INCOME_ACCESS_<OPERADOR>_<USERNAME>)
  - Captcha   → .env  (var TWOCAPTCHA_API_KEY)

Uso:
    python scripts/scp_income_access.py
    python scripts/scp_income_access.py --operador Betano
    python scripts/scp_income_access.py --headful
    python scripts/scp_income_access.py --operador WilliamHill --headful
"""

from __future__ import annotations

import argparse
import base64
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

import httpx
import pandas as pd
from playwright.sync_api import BrowserContext, Page, sync_playwright

# ── Caminhos do projecto ────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parents[1]
CONFIG_XLSX = ROOT / "config" / "logins.xlsx"
DATA_DIR = ROOT / "data"
REPORT_DIR = ROOT / "report"
LOGS_DIR = ROOT / "logs"


def _load_env() -> None:
    """
    Carrega o .env suportando chaves entre aspas ("KEY"="VALUE"),
    que o python-dotenv padrão não strip nas chaves.
    """
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
PLATFORM_SLUG = "INCOME_ACCESS"
ACTIVE_FLAGS = {"1", "4", "5"}
CAPTCHA_FLAGS = {"4", "5"}

# O report está numa SPA Angular — o iframe clássico serve o conteúdo real
# O URL do iframe é o mesmo path sem o prefixo /portal/#
REPORT_PATH = "/reporting/earnings_report.asp"

DEFAULT_TIMEOUT = int(os.getenv("DEFAULT_TIMEOUT", "20000"))
SLOW_MO = int(os.getenv("SLOW_MO", "0"))
HEADLESS = os.getenv("HEADLESS", "true").lower() == "true"
TWOCAPTCHA_API_KEY = os.getenv("TWOCAPTCHA_API_KEY", "").strip()
# Tempo máximo à espera que o report (kendo-grid) termine de gerar — reports
# grandes (ex: Betano) podem demorar minutos. Poll-based: avança assim que pronto.
REPORT_READY_TIMEOUT = int(os.getenv("REPORT_READY_TIMEOUT", "300000"))


# ── Logger ──────────────────────────────────────────────────────────────────
def get_logger(name: str = "scp_income_access") -> logging.Logger:
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
    return re.sub(r"[^A-Z0-9]", "_", str(val).upper())


def env_key_for(operador: str, username: str) -> str:
    return f"PASS_{PLATFORM_SLUG}_{sanitize(operador)}_{sanitize(username)}"


@dataclass
class Account:
    operador: str
    empresa: str
    username: str
    login_url: str
    file_name: str
    has_captcha: bool

    @property
    def password(self) -> str:
        return os.getenv(env_key_for(self.operador, self.username), "").strip()

    @property
    def base_url(self) -> str:
        p = urlparse(self.login_url)
        return f"{p.scheme}://{p.netloc}"


def load_accounts(operador_filter: str | None = None) -> list[Account]:
    if not CONFIG_XLSX.exists():
        log.error(f"Ficheiro de config não encontrado: {CONFIG_XLSX}")
        return []

    df = pd.read_excel(CONFIG_XLSX, dtype=str)
    df = df[
        (df["Plataforma"].str.strip().str.lower() == "incomeaccess")
        & (df["Active"].str.strip().isin(ACTIVE_FLAGS))
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
        # Força idioma inglês adicionando lang=en ao URL de login
        if "lang=" not in url:
            sep = "&" if "?" in url else "?"
            url = f"{url}{sep}lang=en"
        active_flag = str(row.get("Active", "1")).strip()
        accounts.append(
            Account(
                operador=operador,
                empresa=str(row.get("Empresa", "")).strip(),
                username=username,
                login_url=url,
                file_name=str(row.get("FileName", "")).strip()
                or f"{operador}_{username}",
                has_captcha=active_flag in CAPTCHA_FLAGS,
            )
        )

    # Intercala contas com e sem captcha para não sobrecarregar o serviço de captcha
    with_cap = [a for a in accounts if a.has_captcha]
    without_cap = [a for a in accounts if not a.has_captcha]
    interleaved: list[Account] = []
    for i in range(max(len(with_cap), len(without_cap))):
        if i < len(without_cap):
            interleaved.append(without_cap[i])
        if i < len(with_cap):
            interleaved.append(with_cap[i])
    return interleaved


# ── Captcha (2captcha) ───────────────────────────────────────────────────────
class CaptchaError(Exception):
    pass


def solve_image_captcha(page: Page, img_selector: str, max_wait: int = 120) -> str:
    if not TWOCAPTCHA_API_KEY:
        raise CaptchaError(
            "TWOCAPTCHA_API_KEY não definida — necessária para contas com captcha."
        )
    log.info("Captcha detectado — a resolver via 2captcha...")
    el = page.locator(img_selector)
    img_b64 = base64.b64encode(el.screenshot()).decode()

    resp = httpx.post(
        "http://2captcha.com/in.php",
        data={"key": TWOCAPTCHA_API_KEY, "method": "base64", "body": img_b64, "json": 1},
        timeout=30,
    )
    data = resp.json()
    if data.get("status") != 1:
        raise CaptchaError(f"2captcha recusou: {data.get('error_text', data)}")
    captcha_id = str(data["request"])
    log.info(f"Captcha enviado — ID: {captcha_id}")

    time.sleep(10)
    waited = 0
    while waited < max_wait:
        resp = httpx.get(
            "http://2captcha.com/res.php",
            params={"key": TWOCAPTCHA_API_KEY, "action": "get", "id": captcha_id, "json": 1},
            timeout=15,
        )
        data = resp.json()
        if data.get("status") == 1:
            log.info("Captcha resolvido")
            return str(data["request"])
        if data.get("request") != "CAPCHA_NOT_READY":
            raise CaptchaError(f"Erro 2captcha: {data}")
        time.sleep(5)
        waited += 5

    raise CaptchaError(f"Timeout aguardando captcha ({max_wait}s)")


# ── Login Income Access ───────────────────────────────────────────────────────
SEL_USERNAME = "#username"
SEL_PASSWORD = "#password"
SEL_SUBMIT = "button.btn.btn-primary"
SEL_CAPTCHA_IMG = "img[alt='This Is verification Image']"
SEL_CAPTCHA_INPUT = "#strverifyimg"

_ERROR_SELECTORS = (
    ".alert-danger, .alert.alert-danger, #login-error, .login-error, "
    ".error-message, .validation-summary-errors, .alert-error, #errorMessage"
)


def _login_error_text(page: Page) -> str:
    try:
        els = page.locator(_ERROR_SELECTORS)
        texts = []
        for i in range(els.count()):
            el = els.nth(i)
            try:
                if el.is_visible():
                    t = (el.text_content() or "").strip()
                    if t:
                        texts.append(t)
            except Exception:
                pass
        return " | ".join(texts).strip()
    except Exception:
        return ""


def _fill_and_submit(page: Page, acc: Account) -> None:
    """Preenche username/password, resolve captcha se necessário, e clica submit."""
    page.wait_for_selector(SEL_USERNAME, timeout=DEFAULT_TIMEOUT)
    page.fill(SEL_USERNAME, acc.username)
    page.fill(SEL_PASSWORD, acc.password)

    if acc.has_captcha:
        try:
            page.wait_for_selector(SEL_CAPTCHA_IMG, timeout=8000)
        except Exception:
            # Página sem captcha visível — continua sem resolver
            log.debug(f"Captcha não encontrado em {acc.base_url} — a continuar sem resolver")
            page.click(SEL_SUBMIT)
            return
        try:
            solution = solve_image_captcha(page, SEL_CAPTCHA_IMG)
            # Verifica se a imagem do captcha ainda é a mesma (página pode ter actualizado)
            page.wait_for_selector(SEL_CAPTCHA_IMG, state="visible", timeout=3000)
            page.fill(SEL_CAPTCHA_INPUT, solution)
        except CaptchaError as e:
            raise CaptchaError(str(e))
        except Exception:
            # Captcha desapareceu enquanto resolvia — a página actualizou
            log.warning("Captcha actualizou durante resolução — a tentar novamente")
            page.wait_for_selector(SEL_CAPTCHA_IMG, timeout=8000)
            solution = solve_image_captcha(page, SEL_CAPTCHA_IMG)
            page.fill(SEL_CAPTCHA_INPUT, solution)

    page.click(SEL_SUBMIT)


def login(page: Page, acc: Account) -> bool:
    log.info(f"A fazer login em {acc.base_url}")
    page.goto(acc.login_url, wait_until="domcontentloaded", timeout=60000)

    try:
        _fill_and_submit(page, acc)
    except CaptchaError as e:
        log.error(f"Falha no captcha para {acc.operador}/{acc.username}: {e}")
        return False

    page.wait_for_load_state("domcontentloaded", timeout=30000)

    # Poll até 45s: sucesso (área autenticada) vs erro explícito vs novo captcha
    auth_sel = "a[href*='report'], a[href*='Report'], a[href*='logout'], a[href*='commission']"
    deadline = time.monotonic() + 45
    while time.monotonic() < deadline:
        if page.locator(auth_sel).count() > 0:
            log.info("Login OK")
            return True

        # Página voltou ao form de login com novo captcha → re-submete
        if (
            page.locator(SEL_USERNAME).count() > 0
            and page.locator(SEL_USERNAME).is_visible()
        ):
            if page.locator(SEL_CAPTCHA_IMG).count() > 0:
                log.warning("Novo captcha apareceu após submit — a resolver novamente")
                try:
                    _fill_and_submit(page, acc)
                    page.wait_for_load_state("domcontentloaded", timeout=30000)
                    continue
                except CaptchaError as e:
                    log.error(f"Falha no captcha (retry) para {acc.operador}/{acc.username}: {e}")
                    return False
            erro = _login_error_text(page)
            if erro:
                log.error(f"Login falhou para {acc.operador}/{acc.username}: {erro}")
                return False
        elif "login" not in page.url.lower():
            log.info("Login OK")
            return True

        page.wait_for_timeout(1000)

    motivo = _login_error_text(page) or "timeout 45s sem confirmar sessão"
    log.error(f"Login falhou para {acc.operador}/{acc.username}: {motivo}")
    return False




# ── Navegação e extracção do relatório ───────────────────────────────────────
def extract_earnings_report(page: Page, acc: Account) -> pd.DataFrame | None:
    """
    Navega para o Earnings Report (SPA Angular + iframe clássico), define:
      - Start date / End date = primeiro e último dia do mês anterior
      - Merchant = All Merchants (kendo-combobox id="MerchantId")
    Clica em "Generate Report" e exporta o CSV UTF-8.
    """
    report_url = acc.base_url + REPORT_PATH
    page.goto(report_url, wait_until="domcontentloaded", timeout=60000)
    log.info(f"A navegar para: {report_url}")

    # Sessão expirou → voltou ao login
    if page.locator(SEL_USERNAME).count() > 0 and page.locator(SEL_USERNAME).is_visible():
        log.warning("Sessão expirou — a tentar novo login...")
        if not login(page, acc):
            return None
        page.goto(report_url, wait_until="domcontentloaded", timeout=60000)

    # Aguarda o web component Angular renderizar os campos Kendo
    try:
        page.wait_for_selector("#DatePeriod input.k-input-inner", state="visible", timeout=DEFAULT_TIMEOUT)
        page.wait_for_timeout(1000)
    except Exception:
        page.wait_for_timeout(6000)

    # ── Selecciona período = mês anterior via dropdown DatePeriod (4ª opção) ──
    _select_period_dropdown(page)

    # ── Merchant = All Merchants via kendo-combobox id="MerchantId" ──────────
    _set_all_merchants_kendo(page)

    # ── Checkboxes Merchant/Member (apenas para TipsterpageAR) ──────────────
    if acc.username.lower() == "tipsterpagear":
        try:
            cb = page.locator("#cb-2-merchant")
            if cb.count() > 0 and not cb.is_checked():
                page.locator("label[for='cb-2-merchant']").click()
                log.debug("Checkbox 'Merchant' marcada")
        except Exception as e:
            log.debug(f"Checkbox Merchant: {e}")
        try:
            cb = page.locator("#cb-3-member")
            if cb.count() > 0 and cb.is_checked():
                page.locator("label[for='cb-3-member']").click()
                log.debug("Checkbox 'Member' desmarcada")
        except Exception as e:
            log.debug(f"Checkbox Member: {e}")

    # ── Generate Report ───────────────────────────────────────────────────────
    try:
        btn = page.locator("button.button--accept, button[type='submit']").first
        btn.wait_for(state="visible", timeout=8000)
        btn.click()
        log.debug("Generate Report clicado")
    except Exception as e:
        log.warning(f"Botão Generate Report não encontrado: {e}")

    # Aguarda o kendo-grid terminar de gerar o report (header + linhas)
    if not _wait_report_ready(page):
        log.warning(
            f"Report não ficou pronto a tempo para {acc.operador}/{acc.username} "
            f"— a saltar (sem exportar dados parciais)"
        )
        return None

    # ── Export → CSV UTF-8 ────────────────────────────────────────────────────
    df = _export_csv_utf8(page, acc)
    if df is None:
        log.warning(f"Export falhou para {acc.operador}/{acc.username}")
        return None
    if df.empty:
        log.info(f"Sem dados no período para {acc.operador}/{acc.username} — a inserir linha com zeros")
        df = _zero_row_df(list(df.columns))
    return df


def _wait_report_ready(page: Page) -> bool:
    """
    Aguarda (até REPORT_READY_TIMEOUT) que o kendo-grid termine de gerar o report.
    Poll-based: devolve assim que o header do grid estiver visível e os dados
    (ou a indicação de 'sem registos') tiverem carregado. Devolve False se esgotar
    sem o header aparecer — nesse caso o chamador salta a conta em vez de exportar
    dados parciais/corrompidos.
    """
    deadline = time.monotonic() + REPORT_READY_TIMEOUT / 1000.0

    # 1) Header do grid — sinal de que o report renderizou
    try:
        remaining = max(1000, int((deadline - time.monotonic()) * 1000))
        page.wait_for_selector("th[kendogridlogicalcell]", state="visible", timeout=remaining)
    except Exception:
        return False

    # 2) Máscara de loading do Kendo desaparece (se existir)
    try:
        page.wait_for_selector(".k-loading-mask, .k-loader", state="hidden", timeout=10000)
    except Exception:
        pass

    # 3) Aguarda células de dados OU indicação de 'sem registos'
    while time.monotonic() < deadline:
        has_data = page.locator("td[kendogridlogicalcell]").count() > 0
        no_records = page.locator(".k-grid-norecords, .k-no-data").count() > 0
        if has_data:
            log.debug("Report carregado (dados prontos)")
            page.wait_for_timeout(500)  # folga para render completo
            return True
        if no_records:
            log.debug("Report carregado (sem registos)")
            return True
        page.wait_for_timeout(1000)

    # Header apareceu mas dados não confirmados — prossegue mesmo assim
    log.debug("Report: header presente mas dados não confirmados — a prosseguir")
    return True


def _select_period_dropdown(page: Page) -> None:
    """
    Abre o kendo-combobox #DatePeriod e selecciona a 4ª opção (índice 3),
    que corresponde sempre ao mês anterior (ex: May-2026).
    """
    try:
        # Clica no botão caret para abrir a lista
        caret = page.locator("#DatePeriod button.k-input-button").first
        caret.wait_for(state="visible", timeout=8000)
        caret.click()
        page.wait_for_timeout(600)

        # Aguarda as opções aparecerem
        items = page.locator("kendo-popup li.k-list-item")
        items.nth(0).wait_for(state="visible", timeout=8000)

        # 4ª opção = índice 3 = mês anterior
        item = items.nth(3)
        label = (item.text_content() or "").strip()
        item.click()
        log.debug(f"Período seleccionado: {label!r}")
    except Exception as e:
        log.warning(f"Não foi possível seleccionar o período: {e}")


def _set_all_merchants_kendo(page: Page) -> None:
    """
    Define Merchant = All Merchants no kendo-combobox id="MerchantId".
    Limpa o campo e digita "All" para filtrar, depois selecciona a primeira opção.
    """
    try:
        inp = page.locator("#MerchantId input.k-input-inner").first
        inp.wait_for(state="visible", timeout=5000)
        inp.click(click_count=3)
        inp.fill("")
        inp.type("All", delay=80)
        page.wait_for_timeout(800)
        # Selecciona a primeira opção da lista
        option = page.locator("kendo-popup li.k-list-item, ul.k-list-ul li").first
        if option.count() > 0:
            option.click()
            log.debug("Merchant: 'All Merchants' seleccionado")
        else:
            # Fallback: se não apareceu lista, apaga e deixa vazio (= all)
            inp.click(click_count=3)
            inp.fill("")
            inp.press("Tab")
            log.debug("Merchant: campo limpo (fallback = all)")
    except Exception as e:
        log.warning(f"Não foi possível definir Merchant: {e}")


def _export_csv_utf8(page: Page, acc: Account) -> pd.DataFrame | None:
    """
    Clica no botão Export e selecciona a opção CSV UTF-8.
    O botão Export usa kendo (<span class='k-button-text'>Export</span>).
    """
    try:
        # Localiza o botão Export pelo texto
        export_btn = page.get_by_text("Export", exact=True).first
        export_btn.wait_for(state="visible", timeout=8000)
        export_btn.click()
        page.wait_for_timeout(800)

        # Procura opção "CSV (UTF-8)" no dropdown/popup que abre
        csv_option = page.get_by_text("CSV (UTF-8)", exact=False).first
        if csv_option.count() == 0:
            # Fallback: qualquer opção que mencione csv
            csv_option = page.locator("[class*='export'] li, kendo-popup li").filter(
                has_text="CSV"
            ).first

        DATA_DIR.mkdir(parents=True, exist_ok=True)
        tmp_path = DATA_DIR / f"_tmp_{_safe_name(acc.file_name)}.csv"

        with page.expect_download(timeout=60000) as dl_info:
            csv_option.click()

        dl_info.value.save_as(tmp_path)
        log.debug(f"CSV descarregado: {tmp_path.name}")
        df = pd.read_csv(tmp_path, dtype=str, encoding="utf-8-sig")
        tmp_path.unlink(missing_ok=True)
        return df
    except Exception as e:
        log.debug(f"Export CSV UTF-8 falhou ({e}) — a tentar leitura HTML")
        return None


def _zero_row_df(columns: list[str]) -> pd.DataFrame:
    """Cria um DataFrame de uma linha com zeros para contas sem dados no período."""
    NON_NUMERIC = {"rowid", "currency symbol", "affiliate id", "username", "country"}
    row = {}
    for col in columns:
        row[col] = "1" if col.lower() == "rowid" else ("0" if col.lower() not in NON_NUMERIC else "")
    return pd.DataFrame([row])


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


# ── Padronização ──────────────────────────────────────────────────────────────

# Mapeamento das colunas do report Income Access para nomes canónicos.
#
# Apesar de forçarmos lang=en na URL de login, o Income Access guarda a
# preferência de idioma por conta no servidor, pelo que alguns relatórios
# chegam em Português ou Espanhol. Para garantir que os campos ficam sempre
# padronizados independentemente do idioma do report, mapeamos aqui os
# cabeçalhos EN/PT/ES (e variantes por merchant) para o mesmo nome canónico.
# As chaves estão em minúsculas — o lookup é feito com c.lower().
IA_COLUMN_MAP = {
    # ── rowid ────────────────────────────────────────────────────────────────
    "rowid": "rowid",
    # ── currency ─────────────────────────────────────────────────────────────
    "currency symbol": "currency",
    "currencysymbol": "currency",
    "símbolo da moeda": "currency",  # PT
    # ── total records ────────────────────────────────────────────────────────
    "total records": "total_records",
    "totalrecords": "total_records",
    # ── affiliate id ─────────────────────────────────────────────────────────
    "affiliate id": "affiliate_id",
    "id do afiliado": "affiliate_id",          # PT
    "identidad del afiliado": "affiliate_id",  # ES
    # ── affiliate username ───────────────────────────────────────────────────
    "username": "affiliate_username",
    "nome de usuário": "affiliate_username",   # PT
    "nombre de usuario": "affiliate_username",  # ES
    # ── country ──────────────────────────────────────────────────────────────
    "country": "country",
    "país": "country",  # PT
    # ── impressions ──────────────────────────────────────────────────────────
    "impressions": "impressions",
    "impressões": "impressions",  # PT
    "impresiones": "impressions",  # ES
    # ── clicks ───────────────────────────────────────────────────────────────
    "clicks": "clicks",
    "cliques": "clicks",  # PT
    # ── click-through ratio ──────────────────────────────────────────────────
    "click-through ratio": "click_through_ratio",
    "porcentagem de cliques": "click_through_ratio",        # PT
    "porcentaje de 'click-through'": "click_through_ratio",  # ES
    # ── registrations ────────────────────────────────────────────────────────
    "registrations": "registrations",
    "registros": "registrations",  # PT
    "descargas": "registrations",   # ES (Descargas/downloads ocupa este slot)
    # ── registration ratio ───────────────────────────────────────────────────
    "registration ratio": "registration_ratio",
    "porcentagem de registros": "registration_ratio",  # PT
    "porcentaje de 'descargas'": "registration_ratio",  # ES
    # ── deposits ─────────────────────────────────────────────────────────────
    "deposits": "deposits",
    "depósitos": "deposits",  # PT
    "depositos": "deposits",   # ES
    # ── net revenue ──────────────────────────────────────────────────────────
    "net revenue": "net_revenue",
    "netrev": "net_revenue",
    "receita liquida": "net_revenue",  # PT
    "ingresos de red": "net_revenue",   # ES
    # ── gross revenue ────────────────────────────────────────────────────────
    "gross revenue": "gross_revenue",
    "rendimento bruto": "gross_revenue",  # PT
    "ingresos brutos": "gross_revenue",    # ES
    # ── total bets/hands ─────────────────────────────────────────────────────
    "total bets/hands": "total_bets_hands",
    "total de apuestas / manos": "total_bets_hands",  # ES
    # ── stake ────────────────────────────────────────────────────────────────
    "stake": "stake",
    "apuestas": "stake",  # ES
    # ── net points ───────────────────────────────────────────────────────────
    "net points": "net_points",
    # ── % commission ─────────────────────────────────────────────────────────
    "% commission": "pct_commission",
    "%commission": "pct_commission",
    "% de comissão": "pct_commission",       # PT
    "porcentaje de comisión": "pct_commission",  # ES
    # ── cpa commission ───────────────────────────────────────────────────────
    "cpa commission": "cpa_commission",
    "comissão cpa": "cpa_commission",                  # PT
    "comisión coste de adquisición": "cpa_commission",  # ES
    # ── cpa count ────────────────────────────────────────────────────────────
    "cpa count": "cpa_count",
    "conta cpa": "cpa_count",  # PT
    # ── referral commission ──────────────────────────────────────────────────
    "referral commission": "referral_commission",
    "comissão de referentes": "referral_commission",  # PT
    "comisión de referido": "referral_commission",     # ES
    # ── total commission ─────────────────────────────────────────────────────
    "total commission": "total_commission",
    "comissão total": "total_commission",  # PT
    "comisión total": "total_commission",   # ES
    # ── bonus ────────────────────────────────────────────────────────────────
    "bonus": "bonus",
    "bônus": "bonus",  # PT
    # ── chargebacks ──────────────────────────────────────────────────────────
    "chargebacks": "chargebacks",
    "estorno": "chargebacks",                                # PT
    "chargebacks 'reembolsos fraudulentos'": "chargebacks",  # ES
    # ── outras variantes EN por merchant ─────────────────────────────────────
    "installs": "installs",
    "memberid": "member_id",
    "wagers": "wagers",
    "commadminfee": "comm_admin_fee",
    "adjustments": "adjustments",
    "costs": "costs",
    "new account ratio": "new_account_ratio",
    "new depositing acc count": "new_depositing_acc_count",
    "new active acc count": "new_active_acc_count",
    "first deposit count": "first_deposit_count",
    "active accounts": "active_accounts",
    "active days": "active_days",
    "new acc purchases": "new_acc_purchases",
    "depositing accounts": "depositing_accounts",
    "wagering accounts": "wagering_accounts",
    "average active days": "average_active_days",
    "gross / player": "gross_per_player",
    "net / player": "net_per_player",
    # ── merchant (tratado em standardize) ────────────────────────────────────
    "merchant": "merchant",
}


def standardize(df: pd.DataFrame, acc: Account) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Remove linha de totais
    if len(df.columns) > 0:
        first = df.columns[0]
        df = df[
            ~df[first].astype(str).str.strip().str.lower().str.startswith("total")
        ].reset_index(drop=True)

    # Remove colunas sem nome (pandas numera-as 0, 1, 2... ao ler CSV com colunas vazias)
    unnamed = [c for c in df.columns if str(c).strip().isdigit()]
    if unnamed:
        df.drop(columns=unnamed, inplace=True)

    # Normaliza nomes de colunas para os canónicos
    df.rename(
        columns={c: IA_COLUMN_MAP.get(c.lower(), c) for c in df.columns},
        inplace=True,
    )

    # Funde colunas que colidiram no mesmo nome canónico (ex: report que traz
    # tanto "Net Revenue" como a variante "NetRev"): mantém o 1.º valor não-vazio.
    if df.columns.duplicated().any():
        df = df.T.groupby(level=0, sort=False).first().T

    from datetime import date
    _today = date.today().replace(day=1)
    _prev = (_today - __import__("datetime").timedelta(days=1)).replace(day=1)
    month_label = _prev.strftime("%Y-%m")

    # Se o report tiver coluna "merchant" com valores, combina com o username
    # ex: merchant="Betano AR CABA", username="TipsterpageAR" → "CABA_TipsterpageAR"
    if "merchant" in df.columns and df["merchant"].notna().any():
        account_username_val = df["merchant"].apply(
            lambda m: f"{str(m).strip().split()[-1]}_{acc.username}"
            if pd.notna(m) and str(m).strip()
            else acc.username
        )
        df.drop(columns=["merchant"], inplace=True)
    else:
        account_username_val = acc.username

    meta = {
        "operador": acc.operador,
        "empresa": acc.empresa,
        "plataforma": "IncomeAccess",
        "account_username": account_username_val,
        "month": month_label,
        "extracted_at": datetime.now().isoformat(timespec="seconds"),
    }
    for i, (k, v) in enumerate(meta.items()):
        df.insert(i, k, v)
    return df


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
    log.info(f"Guardado {len(df)} linhas -> {path.name}")
    return path


def save_combined(frames: list[pd.DataFrame]) -> Path | None:
    if not frames:
        log.warning("Nenhum dado para consolidar — data/income_access.csv não gerado")
        return None
    new_data = pd.concat(frames, ignore_index=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    path = REPORT_DIR / "income_access.csv"
    merged = _upsert_csv(path, new_data, ["operador", "account_username", "month"])
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

    if acc.has_captcha and not TWOCAPTCHA_API_KEY:
        log.error(
            f"{acc.operador}/{acc.username} requer captcha mas TWOCAPTCHA_API_KEY não está definida"
        )
        return None

    log.info(f"── IncomeAccess / {acc.operador} / {acc.username}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=SLOW_MO)
        context: BrowserContext = browser.new_context(accept_downloads=True)
        page: Page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)
        try:
            if not login(page, acc):
                return None
            df = extract_earnings_report(page, acc)
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
        description="Extração mensal de comissões Income Access",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--operador", help="Processa apenas este operador (ex: Betano)")
    parser.add_argument("--headful", action="store_true", help="Mostra o browser")
    args = parser.parse_args()

    headless = HEADLESS and not args.headful

    log.info("=" * 60)
    log.info(f"Inicio -- {datetime.now():%Y-%m-%d %H:%M:%S}")

    accounts = load_accounts(operador_filter=args.operador)
    if not accounts:
        log.warning("Nenhum operador activo encontrado -- a terminar")
        sys.exit(0)

    captcha_count = sum(1 for a in accounts if a.has_captcha)
    log.info(
        f"{len(accounts)} operador(es) a processar "
        f"({captcha_count} com captcha)"
    )

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
