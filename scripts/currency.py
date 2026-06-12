import ast
import csv
import datetime
import json
import os
from pathlib import Path

import requests

try:  # funciona tanto em execução directa quanto importado como scripts.*
    import common
except ImportError:
    from scripts import common

# Carrega variáveis de ambiente do arquivo .env na raiz do projeto
common.load_env()

# Caminhos absolutos (independentes do diretório de trabalho).
JSON_PATH = str(common.DATA_DIR / "currency.json")
CSV_PATH = str(common.DATA_DIR / "currency_rates.csv")

log = common.get_logger("currency")


# type hint indica o que a função deve retornar (recebe a url como uma str e retorna uma lista)
def extract_currency(
    url: str,
) -> list:
    response = requests.get(
        url
    )  # requisição HTTP do tipe GET para a URL da API, armazenando a resposta em 'response'
    data = response.json()  # converte a resposta para um dicionário Python

    # se o status da resposta for diferente de 200, exibe um erro e retorna uma lista vazia
    if response.status_code != 200:
        log.error(f"Error fetching data from: {response.status_code}")
        return []

    # se não houver dados na resposta, exibe um erro e retorna uma lista vazia
    if not data:
        log.error("No data received from the API")
        return []

    # caminho onde será salvo o arquivo
    output_path = JSON_PATH
    # caminho do diretório onde será salvo o arquivo
    output_dir = Path(output_path).parent
    # cria o diretório se ele não existir
    output_dir.mkdir(parents=True, exist_ok=True)

    # abre o arquivo em modo de escrita
    with open(output_path, "w", encoding="utf-8") as f:
        # escreve os dados no arquivo
        json.dump(data, f, ensure_ascii=False, indent=4)

    log.info(f"Data saved to {output_path}")
    return data


def save_currency(
    json_path: str = JSON_PATH,
    target_currencies: list[str] | None = None,
    csv_path: str = CSV_PATH,
) -> None:
    """Append selected currency rates to a CSV.
    The CSV will contain: time_last_update_utc, base_code, currency_code, rate.
    If the file does not exist it is created with a header row.
    """
    # carrega os dados do JSON
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        log.error(f"JSON file not found: {json_path}")
        return

    # Se não houver lista de moedas, tenta ler do .env
    if target_currencies is None:
        env_val = os.getenv("TARGET_CURRENCIES")
        if env_val:
            try:
                target_currencies = ast.literal_eval(env_val)
            except Exception as e:
                log.error(f"Failed to parse TARGET_CURRENCIES from .env: {e}")
                target_currencies = []
        else:
            target_currencies = []

    raw_timestamp = data.get("time_last_update_utc")
    base = data.get("base_code")
    rates = data.get("conversion_rates", {})

    # Converte data para formato yyyy-MM-dd
    try:
        dt = datetime.datetime.strptime(raw_timestamp, "%a, %d %b %Y %H:%M:%S %z")
        timestamp = dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        log.error(f"Failed to parse timestamp '{raw_timestamp}': {e}")
        timestamp = raw_timestamp
        dt = None

    # Mês de referência = mês anterior à data da cotação (período do relatório).
    # A cotação é puxada no início do mês seguinte ao período extraído pelos scrapers.
    if dt is not None:
        first_of_month = dt.replace(day=1)
        prev = (first_of_month - datetime.timedelta(days=1)).replace(day=1)
        month = prev.strftime("%Y-%m")
    else:
        month = ""

    # Se ainda não houver moedas (lista vazia), grava todas as disponíveis
    if not target_currencies:
        target_currencies = list(rates.keys())

    rows = []
    for cur in target_currencies:
        rate = rates.get(cur)
        if rate is None:
            log.warning(f"Currency {cur} not found in conversion_rates")
            continue
        rows.append([month, timestamp, base, cur, rate])

    # Acumula cotações: dedup por (time_last_update_utc, currency). Assim várias
    # cotações do mesmo mês coexistem e o build_union escolhe a mais antiga do mês.
    existing_pairs = set()
    if Path(csv_path).exists():
        with open(csv_path, newline="", encoding="utf-8") as csvfile_r:
            reader = csv.reader(csvfile_r)
            # pula cabeçalho
            next(reader, None)
            for row in reader:
                if len(row) >= 5:
                    existing_pairs.add((row[1], row[3]))  # (time_last_update_utc, currency)

    # Filtra linhas que ainda não estão no CSV
    filtered_rows = [row for row in rows if (row[1], row[3]) not in existing_pairs]

    if not filtered_rows:
        log.info("All rows already present in CSV; nothing to append")
        return

    rows = filtered_rows
    file_exists = Path(csv_path).exists()
    with open(csv_path, "a", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        if not file_exists:
            writer.writerow(["month", "time_last_update_utc", "base_code", "currency", "rate"])
        writer.writerows(rows)
    log.info(f"Appended {len(rows)} rows to {csv_path}")


if __name__ == "__main__":
    # 1) Baixa/atualiza o JSON a partir da API (se URL_CURRENCY_API estiver no .env).
    #    Sem URL, usa o data/currency.json já existente.
    url = os.getenv("URL_CURRENCY_API")
    if url and "<chave>" not in url:
        extract_currency(url)
    else:
        log.info("URL_CURRENCY_API não configurada — a usar o currency.json existente")

    # 2) Grava as moedas-alvo (TARGET_CURRENCIES no .env) ou todas as disponíveis.
    save_currency(target_currencies=None)
