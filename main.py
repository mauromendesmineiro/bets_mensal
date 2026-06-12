"""
main.py — Orquestrador da extração mensal.

Pipeline (nesta ordem):
  1. currency.py    — actualiza as taxas de câmbio (data/currency_rates.csv)
  2. scp_*.py       — corre os 4 scrapers de plataforma EM PARALELO, cada um no seu
                      próprio processo (a API síncrona do Playwright não é segura
                      entre threads). Dentro de cada plataforma as contas são
                      processadas sequencialmente.
  3. gsheets.py     — actualiza a coluna 'resta' do config/ref.xlsx a partir da
                      planilha do Google (--pull-ref) e valida que toda conta da
                      planilha existe no logins.xlsx (--validate-logins, grava
                      config/contas_faltantes.csv). Antes do build_union.
  4. build_union.py — consolida os report/*.csv em report/union_data.csv

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
import csv
import os
import re
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
GSHEETS_SCRIPT = "gsheets.py"
UNION_SCRIPT = "build_union.py"


def run_step(
    name: str,
    script: str,
    extra_args: list[str] | None = None,
    timeout: float | None = None,
    capture: bool = False,
) -> tuple[str, int, float, str]:
    """Corre um script como subprocesso. Devolve (nome, rc, duração_s, output).

    - `timeout`: segundos; se excedido, o processo é morto e rc = -1 (TIMEOUT).
    - `capture`: True para capturar stdout/stderr (usado no paralelo, para imprimir
      o log em bloco no fim e evitar intercalação). False faz streaming ao vivo.
    """
    cmd = [sys.executable, str(SCRIPTS_DIR / script), *(extra_args or [])]
    start = time.monotonic()
    if not capture:
        print(f"[{name}] arranque → {' '.join(cmd)}", flush=True)
    try:
        proc = subprocess.run(
            cmd, cwd=str(ROOT), timeout=timeout,
            capture_output=capture, text=capture, encoding="utf-8", errors="replace",
        )
        rc = proc.returncode
        out = ((proc.stdout or "") + (proc.stderr or "")) if capture else ""
    except subprocess.TimeoutExpired as e:
        rc = -1
        out = ""
        if capture:
            so = e.stdout.decode("utf-8", "replace") if isinstance(e.stdout, bytes) else (e.stdout or "")
            se = e.stderr.decode("utf-8", "replace") if isinstance(e.stderr, bytes) else (e.stderr or "")
            out = so + se
    dur = time.monotonic() - start
    if rc == 0:
        estado = "OK"
    elif rc == -1:
        estado = f"TIMEOUT (>{timeout:.0f}s)"
    else:
        estado = f"ERRO (rc={rc})"
    if not capture:
        print(f"[{name}] terminou em {_fmt_hms(dur)} — {estado}", flush=True)
    return name, rc, dur, out


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
        "--id", nargs="+", dest="ids",
        help="Filtra por Id de conta do logins.xlsx (passado a cada scraper)",
    )
    parser.add_argument(
        "--headful", action="store_true", help="Mostra os browsers (passa a cada scraper)"
    )
    parser.add_argument(
        "--max-workers", type=int, default=4,
        help="Nº máximo de plataformas em paralelo (por omissão: 4)",
    )
    parser.add_argument(
        "--timeout", type=float, default=1800.0,
        help="Timeout por etapa em segundos (por omissão: 1800 = 30 min)",
    )
    parser.add_argument(
        "--month",
        help="Reprocessa um mês específico (YYYY-MM). Por omissão: mês anterior.",
    )
    parser.add_argument("--no-currency", action="store_true", help="Salta a etapa de câmbio")
    parser.add_argument("--no-ref", action="store_true", help="Salta a atualização do ref via Google Sheets")
    parser.add_argument("--no-union", action="store_true", help="Salta a consolidação final")
    args = parser.parse_args()

    # Mês alvo: validado e exportado para os subprocessos via TARGET_MONTH.
    if args.month:
        if not re.fullmatch(r"\d{4}-\d{2}", args.month):
            parser.error("--month deve estar no formato YYYY-MM (ex.: 2026-05)")
        os.environ["TARGET_MONTH"] = args.month

    selected = args.only or list(SCRAPERS)
    extra: list[str] = []
    if args.operador:
        extra += ["--operador", args.operador]
    if args.ids:
        extra += ["--id", *args.ids]
    if args.headful:
        extra.append("--headful")

    mes_txt = args.month or "mês anterior"
    print("=" * 60)
    print(f"Pipeline — início {datetime.now():%Y-%m-%d %H:%M:%S}  |  período: {mes_txt}")
    print(f"Plataformas: {', '.join(selected)}  |  max paralelo: {args.max_workers}")
    print("=" * 60)

    inicio = time.monotonic()
    resultados: list[tuple[str, int, float]] = []

    # ── 1) Câmbio ───────────────────────────────────────────────────────────
    if args.no_currency:
        print("[câmbio] saltado (--no-currency)")
    else:
        resultados.append(run_step("câmbio", CURRENCY_SCRIPT, timeout=args.timeout))

    # ── 2) Scrapers em paralelo (saída capturada e impressa em bloco) ───────
    print("-" * 60)
    print(f"A correr {len(selected)} scraper(s) em paralelo… (timeout {args.timeout:.0f}s cada)")
    paral: dict[str, tuple[str, int, float, str]] = {}
    with ThreadPoolExecutor(max_workers=args.max_workers) as pool:
        futuros = {
            pool.submit(run_step, name, SCRAPERS[name], extra, args.timeout, True): name
            for name in selected
        }
        for fut in as_completed(futuros):
            r = fut.result()
            paral[r[0]] = r
    # Imprime o log de cada scraper em bloco, em ordem determinística.
    for name in selected:
        nm, rc, dur, out = paral[name]
        estado = "OK" if rc == 0 else ("TIMEOUT" if rc == -1 else f"ERRO(rc={rc})")
        print(f"\n────── [{nm}] {_fmt_hms(dur)} — {estado} ──────")
        if out.strip():
            print(out.rstrip())
        resultados.append((nm, rc, dur, out))

    # ── 3) Atualização do ref via Google Sheets (antes da union) ────────────
    print("-" * 60)
    if args.no_ref:
        print("[ref] saltado (--no-ref)")
    else:
        resultados.append(
            run_step("ref", GSHEETS_SCRIPT, ["--pull-ref", "--validate-logins"], timeout=args.timeout)
        )

    # ── 4) Consolidação ─────────────────────────────────────────────────────
    print("-" * 60)
    if args.no_union:
        print("[union] saltado (--no-union)")
    else:
        resultados.append(run_step("union", UNION_SCRIPT, timeout=args.timeout))

    total = time.monotonic() - inicio
    print("=" * 60)
    ok = [n for n, rc, _, _ in resultados if rc == 0]
    falhou = [n for n, rc, _, _ in resultados if rc != 0]
    for name, rc, dur, _ in resultados:
        marca = "✓" if rc == 0 else "✗"
        print(f"  {marca} {name:<15} {_fmt_hms(dur)}")
    print("-" * 60)
    print(f"Concluído em {_fmt_hms(total)} — {len(ok)} OK, {len(falhou)} com erro")
    if falhou:
        print(f"Com erro: {', '.join(falhou)}")
    print("=" * 60)

    _write_run_summary(resultados, args.month or "")

    sys.exit(1 if falhou else 0)


def _fmt_hms(segundos: float) -> str:
    """Formata uma duração em segundos como hh:mm:ss."""
    total = int(round(segundos))
    h, resto = divmod(total, 3600)
    m, s = divmod(resto, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _write_run_summary(resultados: list[tuple[str, int, float, str]], month: str) -> None:
    """Acrescenta uma linha por etapa em report/run_summary.csv (histórico de execuções)."""
    report_dir = ROOT / "report"
    report_dir.mkdir(parents=True, exist_ok=True)
    path = report_dir / "run_summary.csv"
    ts = datetime.now().isoformat(timespec="seconds")
    novo = not path.exists()
    try:
        with open(path, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if novo:
                w.writerow(["run_at", "month", "step", "returncode", "status", "duration"])
            for name, rc, dur, _ in resultados:
                status = "OK" if rc == 0 else ("TIMEOUT" if rc == -1 else "ERRO")
                w.writerow([ts, month, name, rc, status, _fmt_hms(dur)])
    except Exception as e:
        print(f"  aviso: não foi possível escrever run_summary.csv ({e})")


if __name__ == "__main__":
    main()
