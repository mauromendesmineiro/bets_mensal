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
PLATFORM_SLUG = "NETREFER"


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
        if not username or password.lower() == "nan":
            continue
        key = f"PASS_{PLATFORM_SLUG}_{sanitize(operador)}_{sanitize(username)}"
        if key in seen:
            continue
        seen.add(key)
        lines.append(f"{key}={password}")

    Path(args.output).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"[OK] {len(seen)} passwords escritas em {args.output}")
    print("  Acrescenta o conteúdo ao teu .env e nunca o committes.")


if __name__ == "__main__":
    main()
