# bets_mensal

Extração **mensal** dos ganhos de cada operador de afiliados nas plataformas
**Netrefer**, **Income Access**, **Cellxpert** e **RavenTrack**, sempre referente ao
**mês anterior** à data de execução. Os relatórios de cada plataforma são consolidados
num único dataset (`report/union_data.csv`), enriquecido com a tabela de referência e
publicado no Google Sheets.

## Pipeline

O `main.py` orquestra todo o fluxo. Os 4 scrapers correm **em paralelo** (cada um no
seu próprio processo — a API síncrona do Playwright não é segura entre threads); dentro
de cada plataforma as contas são processadas sequencialmente.

```
1. currency.py     → actualiza taxas de câmbio          (data/currency_rates.csv)
2. scp_*.py  (×4)  → scrapers de plataforma em paralelo (report/<plataforma>.csv)
3. gsheets.py      → puxa o 'ref' do Google Sheets      (config/ref.xlsx, --pull-ref)
                     + valida contas vs logins.xlsx     (config/contas_faltantes.csv)
4. build_union.py  → consolida tudo                     (report/union_data.csv)
5. gsheets.py      → escreve union_data no Sheets       (DatosAutomatizados, append)
```

## Estrutura

```
bets_mensal/
├── main.py                       # orquestrador do pipeline completo
├── config/
│   ├── logins.xlsx               # URLs, usernames, operadores, plataforma e flags (Active)
│   ├── ref.xlsx                  # tabela de referência (operador/empresa/login + 'resta')
│   ├── contas_faltantes.csv      # contas da planilha sem login configurado (gerado)
│   ├── oauth_client.json         # credenciais OAuth Desktop do Google (NÃO committar)
│   └── authorized_user.json      # token OAuth gerado no 1º login (NÃO committar)
├── scripts/
│   ├── common.py                 # utilitários partilhados (env, logger, upsert, moeda, retry…)
│   ├── currency.py               # actualiza taxas de câmbio (ExchangeRate-API)
│   ├── scp_netrefer.py           # scraper Netrefer
│   ├── scp_income_access.py      # scraper Income Access (com captcha via 2captcha)
│   ├── scp_cellxpert.py          # scraper Cellxpert (SPA Angular white-label)
│   ├── scp_raventrack.py         # scraper RavenTrack (Raven5, breakdown por vendor)
│   ├── gsheets.py                # integração Google Sheets (pull ref / push union)
│   ├── build_union.py            # consolida report/*.csv → union_data.csv
│   └── seed_env_passwords.py     # gera as vars PASS_* do .env a partir do xlsx
├── report/                       # CSVs consolidados (ignorados pelo Git)
│   ├── netrefer.csv  income_access.csv  cellxpert.csv  raventrack.csv
│   ├── union_data.csv            # dataset unificado final
│   ├── divergencias_cadastro.csv # linhas do union sem correspondência no ref
│   └── run_summary.csv           # histórico de execuções (por etapa)
├── data/                         # CSVs por conta + câmbio (ignorados pelo Git)
├── tests/                        # pytest (funções puras: common, build_union, gsheets)
├── logs/                         # logs diários por plataforma (ignorados pelo Git)
├── .github/workflows/ci.yml      # CI: uv sync → ruff → pytest
├── .pre-commit-config.yaml       # ruff + ruff-format antes de cada commit
├── .env                          # passwords e config (NÃO committar)
└── .env.example                  # template
```

## Flags Active (logins.xlsx)

| Flag | Significado |
|------|-------------|
| `1`  | Activa, sem captcha |
| `4`  | Activa, com captcha (Betano) |
| `5`  | Activa, com captcha (Betano AR) |
| `0`  | Inactiva — ignorada |
| `3`  | Ignorada (estrutura distinta) |

Cada scraper filtra automaticamente pela coluna `Plataforma`
(`Netrefer`, `IncomeAccess`, `Cellxpert` ou `RavenTrack`).

## As plataformas

| Plataforma | Particularidades |
|---|---|
| **Netrefer** | Login → força idioma EN → `/affiliates/Earnings/MonthlyEarnings`. Download CSV (ou leitura HTML como fallback). `Revenue Share Reward`/`CPA Reward`. |
| **Income Access** | Login com resolução de captcha via **2captcha** (retry automático). Período = mês anterior, Merchant = *All Merchants*. Exporta CSV UTF-8 do kendo-grid (até 5 min). `pct_commission`/`cpa_commission`. |
| **Cellxpert** | SPA Angular servida sob vários domínios white-label (evoaffiliates, blaze.partners, …); já em inglês. Agrega `amount` por `commission_type` (`CPA*` / `Revshare*`). |
| **RavenTrack** | Plataforma Raven5, breakdown por **vendor** (uma linha por `Vendor Name`). Usa o "Export All" para garantir todas as linhas, mas lê a moeda da tabela renderizada. `RevShare Commission`/`CPA Commission`. |

## Escrita no Google Sheets (`DatosAutomatizados`)

Após consolidar o `union_data.csv`, o pipeline **escreve automaticamente** na aba
`DatosAutomatizados` da planilha do Google:

- **Modo:** append (adiciona ao final, não substitui)
- **Decimais:** convertidos para vírgula (ex: `379.71` → `379,71`)
- **Tipos:** números e datas são interpretados corretamente (não como texto)
- **Frequência:** automática a cada execução do `main.py`
- **Flags:**
  - `--no-append` para pular essa etapa
  - `--append-union` para rodar isoladamente

**Alterações manuais:** é possível preencher campos específicos com textos (ex:
`RS_Operador = "Esperar cambio"`) no CSV antes do append; a escrita sobrescreve
a linha inteira.

---

## O dataset unificado (`union_data.csv`)

`build_union.py` normaliza os 4 relatórios num esquema comum:

- **Colunas comuns:** `plataforma, operador, empresa, username, month, currency`
- **Derivadas:** `vendor_name` (só RavenTrack), `rs_operador` (revenue share, numérico),
  `cpa_operador` (CPA, numérico)

Os valores monetários (`R$379.71`, `€1,234.56`, `250477.99`…) são convertidos para número
e a moeda fica em `currency`. O join com `config/ref.xlsx` (case-insensitive, tolera
aliases de operador) enriquece o dataset; divergências de cadastro são reportadas em
`report/divergencias_cadastro.csv`.

## Cadastro: `logins.xlsx` vs `ref.xlsx`

São dois arquivos com **granularidades diferentes e propositais** — não devem ser
unidos, porque a relação entre eles é **1:N**:

| Arquivo | Grão | Responde |
|---|---|---|
| `config/logins.xlsx` | 1 linha = 1 **login** a raspar | *quem eu logo e extraio?* |
| `config/ref.xlsx` | 1 linha = 1 **login × moeda/região** | *como classifico cada valor?* |

Uma única login pode gerar **várias** linhas de referência separadas por moeda/região:
`SportiumBet / Affiliabet` (1 login) → `SPORTIUMBET MEXICO` (MXN) + `SPORTIUMBET PANAMA`
(USD); `Betano / TipsterpageAR` → `BETANO ARGENTINA - CABA` + `…- PBA`. Por isso o
`build_union` desdobra uma raspagem em várias linhas do union conforme a `currency`.

A **fonte de verdade** do cadastro é o **Google Sheets** (4 abas `RS/CPA_*`). A inserção
de contas novas no `logins.xlsx` e no `ref.xlsx` é **manual**.

### Validar contas (`--validate-logins`)

Garante que toda conta presente na planilha está configurada no `logins.xlsx`:

```bash
uv run python scripts/gsheets.py --validate-logins
```

Compara por `(empresa, plataforma, username)` e grava **`config/contas_faltantes.csv`**
com as contas da planilha sem login. Os sub-registos sintéticos de moeda/região
(`affiliabet_panama`, `caba_…`, `pba_…`) são reconhecidos pela login-base e **não**
contam como falta. Roda automaticamente na etapa `ref` do `main.py`.

## Credenciais

- **URLs, usernames, operadores e flags** → `config/logins.xlsx`.
- **Passwords** → **sempre** do `.env`, nunca lidas do Excel.
- **Captcha** → `TWOCAPTCHA_API_KEY` no `.env`.
- **Google Sheets** → OAuth Desktop (`oauth_client.json`); o token é gerado no 1º login.

Convenção das variáveis de password:

```
PASS_NETREFER_<OPERADOR>_<USERNAME>=password
PASS_INCOME_ACCESS_<OPERADOR>_<USERNAME>=password
PASS_CELLXPERT_<OPERADOR>_<USERNAME>=password
PASS_RAVENTRACK_<OPERADOR>_<USERNAME>=password
```

(@, ., espaços e outros caracteres viram `_`; tudo em maiúsculas.)

Para gerar todas de uma vez a partir da coluna `Password` do xlsx:

```bash
uv run python scripts/seed_env_passwords.py   # escreve .env.passwords
# acrescenta o conteúdo de .env.passwords ao teu .env
```

## Setup

```bash
uv sync
uv run playwright install chromium
cp .env.example .env    # preenche passwords, TWOCAPTCHA_API_KEY e config do Google
```

Para o Google Sheets: cria credenciais OAuth "App de computador" no Google Cloud (com a
Google Sheets API activada), guarda o JSON em `config/oauth_client.json` e autoriza no
navegador na 1ª execução com `--pull-ref`.

## Executar

### Pipeline completo

```bash
uv run python main.py                          # pipeline completo
uv run python main.py --only cellxpert raventrack
uv run python main.py --operador Novibet       # filtra o operador nos scrapers
uv run python main.py --id 12 34               # filtra por Id de conta do logins.xlsx
uv run python main.py --month 2026-05          # reprocessa um mês específico (YYYY-MM)
uv run python main.py --headful                # mostra os browsers
uv run python main.py --max-workers 4          # nº máx. de plataformas em paralelo
uv run python main.py --no-currency            # salta a etapa de câmbio
uv run python main.py --no-ref                 # salta a actualização via Google Sheets
uv run python main.py --no-union               # salta a consolidação final
uv run python main.py --no-append              # salta a escrita no Google Sheets
```

Cada execução acrescenta uma linha por etapa em `report/run_summary.csv` (estado e duração em hh:mm:ss).

### Scrapers individuais

```bash
uv run python scripts/scp_netrefer.py
uv run python scripts/scp_income_access.py --operador Betano
uv run python scripts/scp_cellxpert.py --operador BetsAmigo
uv run python scripts/scp_raventrack.py --headful
```

Todos os scrapers aceitam `--operador`, `--id` e `--headful`.

### Etapas avulsas

```bash
uv run python scripts/currency.py                       # actualiza câmbio
uv run python scripts/gsheets.py --pull-ref             # planilha → config/ref.xlsx (merge 'resta')
uv run python scripts/gsheets.py --validate-logins      # planilha → config/contas_faltantes.csv
uv run python scripts/gsheets.py --push-union           # union_data.csv → planilha (substitui tudo)
uv run python scripts/gsheets.py --append-union         # union_data.csv → DatosAutomatizados (append, decimais em vírgula)
uv run python scripts/build_union.py                    # consolida report/*.csv
```

## Variáveis de ambiente

| Variável | Padrão | Descrição |
|----------|--------|-----------|
| `HEADLESS` | `true` | `false` (ou `--headful`) para ver o browser |
| `SLOW_MO` | `0` | ms de atraso entre acções (debug) |
| `DEFAULT_TIMEOUT` | `20000` | timeout geral Playwright (ms) |
| `REPORT_READY_TIMEOUT` | `300000` | tempo máx. à espera do report Income Access (ms) |
| `RETRY_ATTEMPTS` | `3` | tentativas por conta antes de desistir |
| `RETRY_BACKOFF_S` | `10` | backoff entre tentativas (s) |
| `TWOCAPTCHA_API_KEY` | — | chave 2captcha (contas Betano) |
| `CURRENCY_API_KEY` / `URL_CURRENCY_API` | — | ExchangeRate-API |
| `GSHEET_ID` / `GSHEET_REF_TAB` / `GSHEET_UNION_TAB` | — | planilha Google e abas |
| `GOOGLE_OAUTH_CLIENT_JSON` / `GOOGLE_OAUTH_TOKEN_JSON` | — | credenciais/token OAuth |

## Desenvolvimento

```bash
uv run pytest          # testes das funções puras (common, build_union, gsheets)
uv run ruff check      # lint
uv run ruff format     # formatação
```

CI em `.github/workflows/ci.yml` corre `uv sync → ruff check → pytest` em cada push/PR.
O `.pre-commit-config.yaml` corre `ruff` + `ruff-format` antes de cada commit.

## Segurança

- Nenhuma password, chave ou token fica no código — tudo via `.env` / `config/*.json`.
- O `.gitignore` ignora `.env`, `.env.passwords`, `*.xlsx`, `*.csv`, `*.json` e logs.
- `.env.example` é o único ficheiro de ambiente versionado.
- Rotação do token OAuth: apaga `config/authorized_user.json` (regenerado no próximo
  `--pull-ref`) e, se necessário, revoga o acesso em
  https://myaccount.google.com/permissions.
```