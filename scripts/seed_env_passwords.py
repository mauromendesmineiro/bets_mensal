"""
scripts/seed_env_passwords.py

Gera as variáveis PASS_* a partir da coluna Password do config/logins.xlsx,
no formato esperado por scp_netrefer.py:

    PASS_NETREFER_<OPERADOR>_<USERNAME>=<password>

Escreve para .env.passwords (NUNCA committar). Depois copia/acrescenta o
conteúdo ao teu .env. A coluna Password do Excel existe apenas como fonte
desta migração — o scraper lê SEMPRE do ambiente, nunca do xlsx.

Uso:
    python scripts/seed_env_passwords.py
    python scripts/seed_env_passwords.py --output .env.passwords
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CONFIG_XLSX = ROOT / "config" / "logins.xlsx"

PLATFORM_SLUG_MAP = {
    "netrefer": "NETREFER",
    "incomeaccess": "INCOME_ACCESS",
    "cellxpert": "CELLXPERT",
    "raventrack": "RAVENTRACK",
}


def sanitize(val: str) -> str:
    return re.sub(r"[^A-Z0-9]", "_", str(val).upper())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", default=str(ROOT / ".env.passwords"))
    parser.add_argument(
        "--only-active", action="store_true", help="Só linhas com Active == 1"
    )
    args = parser.parse_args()

    df = pd.read_excel(CONFIG_XLSX, dtype=str)
    if args.only_active and "Active" in df.columns:
        df = df[df["Active"].astype(str).str.strip() == "1"].reset_index(drop=True)

    lines = ["# Passwords geradas por seed_env_passwords.py — NÃO committar!", ""]
    seen: set[str] = set()
    for _, row in df.iterrows():
        operador = str(row.get("Operador", "")).strip()
        username = str(row.get("Username", "")).strip()
        password = str(row.get("Password", "")).strip()
        plataforma = str(row.get("Plataforma", "")).strip().lower()
        if not username or password.lower() == "nan":
            continue
        slug = PLATFORM_SLUG_MAP.get(plataforma, sanitize(plataforma))
        key = f"PASS_{slug}_{sanitize(operador)}_{sanitize(username)}"
        if key in seen:
            continue
        seen.add(key)
        lines.append(f'{key}="{password}"')

    Path(args.output).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"[OK] {len(seen)} passwords escritas em {args.output}")
    print("  Acrescenta o conteúdo ao teu .env e nunca o committes.")

    # Higiene: a coluna Password do xlsx é apenas fonte desta migração. Se ainda
    # tiver senhas, recomenda-se limpá-la (o scraper lê SEMPRE do .env).
    com_senha = df["Password"].notna() & (df["Password"].astype(str).str.strip() != "") \
        if "Password" in df.columns else None
    if com_senha is not None and int(com_senha.sum()) > 0:
        print(
            f"  AVISO: a coluna 'Password' do logins.xlsx ainda tem "
            f"{int(com_senha.sum())} valor(es). Após migrar para o .env, "
            f"considera limpá-la (senhas em claro no Excel)."
        )


if __name__ == "__main__":
    main()
