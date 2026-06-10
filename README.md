# bets_mensal

Extração **mensal** dos ganhos de cada operador nas plataformas **Netrefer** e **Income Access**,
referente sempre ao **mês anterior** à data de execução.

## Estrutura

```
bets_mensal/
├── config/
│   └── logins.xlsx                  # URLs, usernames, operadores, plataforma e flags (Active)
├── scripts/
│   ├── scp_netrefer.py              # scraper Netrefer (login → relatório → CSV)
│   ├── scp_income_access.py         # scraper Income Access (login → relatório → CSV)
│   └── seed_env_passwords.py        # gera as vars PASS_* do .env a partir do xlsx
├── data/                            # CSVs gerados (ignorados pelo Git)
│   ├── <Operador>_<User>.csv        # um por operador/conta
│   ├── report_netrefer.csv          # consolidado Netrefer
│   └── income_access.csv            # consolidado Income Access
├── logs/                            # logs diários (ignorados pelo Git)
├── .env                             # passwords e config (NÃO committar)
└── .env.example                     # template
```

## Flags Active (logins.xlsx)

| Flag | Significado |
|------|-------------|
| `1`  | Activa, sem captcha |
| `4`  | Activa, com captcha (Betano) |
| `5`  | Activa, com captcha (Betano AR) |
| `0`  | Inactiva — ignorada |
| `3`  | Ignorada (estrutura distinta) |

Cada script filtra automaticamente pela coluna `Plataforma` (`Netrefer` ou `IncomeAccess`).

## Como funciona

### Netrefer (`scp_netrefer.py`)

1. Lê operadores activos com `Plataforma == Netrefer` de `config/logins.xlsx`.
2. Login e força idioma **EN**.
3. Navega para `/affiliates/Earnings/MonthlyEarnings`.
4. Extrai a tabela (download CSV ou leitura HTML como fallback).
5. Detecta a moeda pelas colunas de valores (€, R$, $, £…).
6. Guarda `data/<Operador>_<User>.csv` com upsert por `(operador, username, month)`.
7. Consolida em `data/report_netrefer.csv`.

### Income Access (`scp_income_access.py`)

1. Lê operadores activos com `Plataforma == IncomeAccess` de `config/logins.xlsx`.
2. A ordem de processamento é intercalada: sem captcha → com captcha → sem captcha…
3. Login (com resolução de captcha via **2captcha** se necessário, com retry automático).
4. Força idioma EN via `?lang=en` no URL de login.
5. Navega para `/reporting/earnings_report.asp`.
6. Selecciona o período = **mês anterior** via dropdown `DatePeriod` (4ª opção).
7. Define Merchant = **All Merchants**.
8. Para `TipsterpageAR`: marca checkbox *Merchant* e desmarca *Member* (obtém breakdown por merchant).
9. Clica **Generate Report** e aguarda o kendo-grid renderizar (até 5 min).
10. Exporta **CSV UTF-8**; se vazio, insere linha com zeros.
11. Guarda `data/<Operador>_<User>.csv` com upsert por `(operador, account_username, month)`.
12. Consolida em `data/income_access.csv`.

## Credenciais

- **URLs, usernames, operadores e flags** → `config/logins.xlsx`.
- **Passwords** → **sempre** do `.env`, nunca lidas do Excel.
- **Captcha** → `TWOCAPTCHA_API_KEY` no `.env`.

Convenção das variáveis de ambiente:

```
PASS_NETREFER_<OPERADOR>_<USERNAME>=password
PASS_INCOME_ACCESS_<OPERADOR>_<USERNAME>=password
TWOCAPTCHA_API_KEY=chave
```

Para gerar todas de uma vez a partir da coluna `Password` do xlsx:

```bash
python scripts/seed_env_passwords.py   # escreve .env.passwords
# acrescenta o conteúdo de .env.passwords ao teu .env
```

## Setup

```bash
uv sync
uv run playwright install chromium
cp .env.example .env    # preenche passwords e TWOCAPTCHA_API_KEY
```

## Executar

```bash
# Netrefer — todos os operadores
uv run python scripts/scp_netrefer.py

# Netrefer — apenas um operador
uv run python scripts/scp_netrefer.py --operador SNAI

# Income Access — todos os operadores
uv run python scripts/scp_income_access.py

# Income Access — apenas um operador
uv run python scripts/scp_income_access.py --operador Betano

# Ver o browser a trabalhar (qualquer script)
uv run python scripts/scp_netrefer.py --headful
uv run python scripts/scp_income_access.py --headful
```

## Variáveis de ambiente opcionais

| Variável | Padrão | Descrição |
|----------|--------|-----------|
| `HEADLESS` | `true` | `false` para ver o browser |
| `SLOW_MO` | `0` | ms de atraso entre acções (debug) |
| `DEFAULT_TIMEOUT` | `20000` | timeout geral Playwright (ms) |
| `REPORT_READY_TIMEOUT` | `300000` | tempo máx. à espera do report Income Access (ms) |

## Segurança

- Nenhuma password ou chave fica no código — tudo via `.env`.
- O `.gitignore` ignora `.env`, `.env.passwords`, `*.xlsx`, `*.csv`, `*.json` e logs.
- `.env.example` é o único ficheiro de ambiente versionado.
