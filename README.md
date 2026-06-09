# bets_mensal

Extração **mensal** dos ganhos de cada operador na plataforma **Netrefer**,
a partir do relatório **Monthly Earnings** (`/affiliates/Earnings/MonthlyEarnings`).

O script está preparado para incluir novos operadores no futuro — basta
acrescentar linhas ao `config/logins.xlsx` e as passwords ao `.env`.

## Estrutura

```
bets_mensal/
├── config/
│   └── logins.xlsx              # URLs, usernames, operadores e flags (Active)
├── scripts/
│   ├── scp_netrefer.py          # scraper principal (login → relatório → CSV)
│   └── seed_env_passwords.py    # gera as vars PASS_* do .env a partir do xlsx
├── data/                        # CSVs gerados (ignorados pelo Git)
│   ├── <Operador>_<User>.csv    # um por operador
│   └── netrefer.csv             # consolidado de todos os operadores
├── .env                         # passwords e config (NÃO committar)
└── .env.example                 # template
```

## Como funciona

1. Lê os operadores **activos** (`Active == 1`) de `config/logins.xlsx`.
2. Para cada operador: login no Netrefer e força o idioma da conta para **EN**.
3. Navega para `/affiliates/Earnings/MonthlyEarnings` (sem seleção de período).
4. Extrai a tabela (download CSV do DataTables, ou leitura do HTML como fallback).
5. Padroniza num `DataFrame` e guarda `data/<Operador>_<User>.csv`.
6. Consolida tudo em `data/netrefer.csv`.

## Credenciais

- **URLs, usernames, operadores e flags** → `config/logins.xlsx`.
- **Passwords** → **sempre** do `.env`, nunca lidas do Excel.

Convenção da variável de ambiente:

```
PASS_NETREFER_<OPERADOR>_<USERNAME>=apassword
```

(`@`, `.`, espaços e outros caracteres viram `_`; tudo em maiúsculas.)
Exemplo: operador `SNAI`, username `Scom1` → `PASS_NETREFER_SNAI_SCOM1`.

Para gerar todas de uma vez a partir da coluna `Password` do xlsx:

```bash
python scripts/seed_env_passwords.py        # escreve .env.passwords
# acrescenta o conteúdo de .env.passwords ao teu .env
```

## Setup

```bash
uv sync
uv run playwright install chromium

cp .env.example .env        # e preenche as passwords
```

## Executar

```bash
# Todos os operadores activos
uv run python scripts/scp_netrefer.py

# Apenas um operador
uv run python scripts/scp_netrefer.py --operador SNAI

# Ver o browser a trabalhar
uv run python scripts/scp_netrefer.py --headful
```

## Segurança

- Nenhuma password ou chave fica no código — tudo via `.env`.
- O `.gitignore` ignora `.env`, `.env.passwords`, `*.xlsx`, `*.csv`, `*.json` e logs.
- `.env.example` é o único ficheiro de ambiente versionado.
