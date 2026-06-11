"""
main.py — Orquestrador dos scrapers mensais.

Corre os 4 scrapers de plataforma em PARALELO, cada um no seu próprio processo
(a API síncrona do Playwright não é segura entre threads, por isso usamos
processos separados — um browser isolado por plataforma). Dentro de cada
plataforma as contas continuam a ser processadas sequencialmente.

Uso:
    python main.py                       # corre as 4 plataformas em paralelo
    python main.py --only cellxpert raventrack
    python main.py --headful             # mostra os browsers
    python main.py --operador Novibet    # filtra o operador em todas as plataformas
    python main.py --max-workers 4       # nº máximo de plataformas em simultâneo
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SCRIPTS_DIR = ROOT / "scripts"

# Plataforma → script. A ordem é apenas a de arranque; correm em paralelo.
SCRAPERS: dict[str, str] = {
    "netrefer": "scp_netrefer.py",
    "income_access": "scp_income_access.py",
    "cellxpert": "scp_cellxpert.py",
    "raventrack": "scp_raventrack.py",
}


def run_scraper(name: str, script: str, extra_args: list[str]) -> tuple[str, int, float]:
    """Corre um scraper como subprocesso. Devolve (nome, returncode, duração_s)."""
    cmd = [sys.executable, str(SCRIPTS_DIR / script), *extra_args]
    start = time.monotonic()
    print(f"[{name}] arranque → {' '.join(cmd)}", flush=True)
    proc = subprocess.run(cmd, cwd=str(ROOT))
    dur = time.monotonic() - start
    estado = "OK" if proc.returncode == 0 else f"ERRO (rc={proc.returncode})"
    print(f"[{name}] terminou em {dur:0.1f}s — {estado}", flush=True)
    return name, proc.returncode, dur


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Orquestrador paralelo dos scrapers mensais",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--only",
        nargs="+",
        choices=sorted(SCRAPERS),
        help="Corre apenas estas plataformas (por omissão: todas)",
    )
    parser.add_argument(
        "--operador", help="Filtra o operador (passado a cada scraper)"
    )
    parser.add_argument(
        "--headful", action="store_true", help="Mostra os browsers (passa a cada scraper)"
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=4,
        help="Nº máximo de plataformas em paralelo (por omissão: 4)",
    )
    args = parser.parse_args()

    selected = args.only or list(SCRAPERS)

    # Argumentos comuns reencaminhados a cada scraper.
    extra: list[str] = []
    if args.operador:
        extra += ["--operador", args.operador]
    if args.headful:
        extra.append("--headful")

    print("=" * 60)
    print(f"Orquestrador — início {datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"Plataformas: {', '.join(selected)}  |  max paralelo: {args.max_workers}")
    print("=" * 60)

    inicio = time.monotonic()
    resultados: list[tuple[str, int, float]] = []
    with ThreadPoolExecutor(max_workers=args.max_workers) as pool:
        futuros = {
            pool.submit(run_scraper, name, SCRAPERS[name], extra): name
            for name in selected
        }
        for fut in as_completed(futuros):
            resultados.append(fut.result())

    total = time.monotonic() - inicio
    print("=" * 60)
    ok = [n for n, rc, _ in resultados if rc == 0]
    falhou = [n for n, rc, _ in resultados if rc != 0]
    for name, rc, dur in sorted(resultados, key=lambda r: r[0]):
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
