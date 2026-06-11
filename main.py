"""
main.py — Orquestrador da extração mensal.

Pipeline (nesta ordem):
  1. currency.py    — actualiza as taxas de câmbio (data/currency_rates.csv)
  2. scp_*.py       — corre os 4 scrapers de plataforma EM PARALELO, cada um no seu
                      próprio processo (a API síncrona do Playwright não é segura
                      entre threads). Dentro de cada plataforma as contas são
                      processadas sequencialmente.
  3. build_union.py — consolida os report/*.csv em report/union_data.csv

Uso:
    python main.py                       # pipeline completo
    python main.py --only cellxpert raventrack
    python main.py --headful             # mostra os browsers
    python main.py --operador Novibet    # filtra o operador nos scrapers
    python main.py --max-workers 4       # nº máximo de plataformas em simultâneo
    python main.py --no-currency         # salta a etapa de câmbio
    python main.py --no-union            # salta a consolidação final
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

# Console em UTF-8 (evita UnicodeEncodeError com →/✓ em terminais cp1252).
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
SCRIPTS_DIR = ROOT / "scripts"

# Plataforma → script. A ordem é apenas a de arranque; correm em paralelo.
SCRAPERS: dict[str, str] = {
    "netrefer": "scp_netrefer.py",
    "income_access": "scp_income_access.py",
    "cellxpert": "scp_cellxpert.py",
    "raventrack": "scp_raventrack.py",
}

CURRENCY_SCRIPT = "currency.py"
UNION_SCRIPT = "build_union.py"


def run_step(name: str, script: str, extra_args: list[str] | None = None) -> tuple[str, int, float]:
    """Corre um script como subprocesso (sequencial). Devolve (nome, rc, duração_s)."""
    cmd = [sys.executable, str(SCRIPTS_DIR / script), *(extra_args or [])]
    start = time.monotonic()
    print(f"[{name}] arranque → {' '.join(cmd)}", flush=True)
    proc = subprocess.run(cmd, cwd=str(ROOT))
    dur = time.monotonic() - start
    estado = "OK" if proc.returncode == 0 else f"ERRO (rc={proc.returncode})"
    print(f"[{name}] terminou em {dur:0.1f}s — {estado}", flush=True)
    return name, proc.returncode, dur


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Orquestrador da extração mensal (câmbio → scrapers → consolidação)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--only", nargs="+", choices=sorted(SCRAPERS),
        help="Corre apenas estas plataformas (por omissão: todas)",
    )
    parser.add_argument("--operador", help="Filtra o operador (passado a cada scraper)")
    parser.add_argument(
        "--headful", action="store_true", help="Mostra os browsers (passa a cada scraper)"
    )
    parser.add_argument(
        "--max-workers", type=int, default=4,
        help="Nº máximo de plataformas em paralelo (por omissão: 4)",
    )
    parser.add_argument("--no-currency", action="store_true", help="Salta a etapa de câmbio")
    parser.add_argument("--no-union", action="store_true", help="Salta a consolidação final")
    args = parser.parse_args()

    selected = args.only or list(SCRAPERS)
    extra: list[str] = []
    if args.operador:
        extra += ["--operador", args.operador]
    if args.headful:
        extra.append("--headful")

    print("=" * 60)
    print(f"Pipeline — início {datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"Plataformas: {', '.join(selected)}  |  max paralelo: {args.max_workers}")
    print("=" * 60)

    inicio = time.monotonic()
    resultados: list[tuple[str, int, float]] = []

    # ── 1) Câmbio ───────────────────────────────────────────────────────────
    if args.no_currency:
        print("[câmbio] saltado (--no-currency)")
    else:
        resultados.append(run_step("câmbio", CURRENCY_SCRIPT))

    # ── 2) Scrapers em paralelo ─────────────────────────────────────────────
    print("-" * 60)
    print(f"A correr {len(selected)} scraper(s) em paralelo…")
    with ThreadPoolExecutor(max_workers=args.max_workers) as pool:
        futuros = {
            pool.submit(run_step, name, SCRAPERS[name], extra): name
            for name in selected
        }
        for fut in as_completed(futuros):
            resultados.append(fut.result())

    # ── 3) Consolidação ─────────────────────────────────────────────────────
    print("-" * 60)
    if args.no_union:
        print("[union] saltado (--no-union)")
    else:
        resultados.append(run_step("union", UNION_SCRIPT))

    total = time.monotonic() - inicio
    print("=" * 60)
    ok = [n for n, rc, _ in resultados if rc == 0]
    falhou = [n for n, rc, _ in resultados if rc != 0]
    for name, rc, dur in resultados:
        marca = "✓" if rc == 0 else "✗"
        print(f"  {marca} {name:<15} {dur:6.1f}s")
    print("-" * 60)
    print(f"Concluído em {total:0.1f}s — {len(ok)} OK, {len(falhou)} com erro")
    if falhou:
        print(f"Com erro: {', '.join(falhou)}")
    print("=" * 60)

    sys.exit(1 if falhou else 0)


if __name__ == "__main__":
    main()
