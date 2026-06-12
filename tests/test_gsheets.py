"""Testes do scripts/gsheets.py que não exigem rede (parsing de aba)."""

import gsheets


class _FakeWorksheet:
    def __init__(self, values):
        self._values = values

    def get_all_values(self):
        return self._values


class _FakeSpreadsheet:
    def __init__(self, mapa):
        self._mapa = mapa

    def worksheet(self, tab):
        return _FakeWorksheet(self._mapa[tab])


def test_read_tab_with_header_linha_nao_inicial():
    # Cabeçalho na linha 3; 2 linhas de "lixo" antes; uma linha vazia no fim.
    values = [
        ["título", "", ""],
        ["", "", ""],
        ["Plataforma", "Usuario Referencia", "Resta"],  # header (linha 3)
        ["Netrefer", "Affiliabet20bet", "No"],
        ["Cellxpert", "AffiliabetBR", "Sí"],
        ["", "", ""],  # linha vazia -> removida
    ]
    sh = _FakeSpreadsheet({"aba": values})
    cols = ["Plataforma", "Usuario Referencia", "Resta"]

    df = gsheets._read_tab_with_header(sh, "aba", 3, cols)

    assert list(df.columns) == cols
    assert len(df) == 2  # linha vazia removida
    assert df.iloc[0]["Usuario Referencia"] == "Affiliabet20bet"
    assert df.iloc[1]["Resta"] == "Sí"


def test_read_tab_with_header_coluna_ausente_fica_vazia():
    values = [
        ["Plataforma", "Resta"],          # header (linha 1)
        ["Netrefer", "No"],
    ]
    sh = _FakeSpreadsheet({"aba": values})
    # Pede uma coluna que não existe na aba -> deve vir vazia, sem erro.
    df = gsheets._read_tab_with_header(sh, "aba", 1, ["Plataforma", "Operador Referencia", "Resta"])
    assert "Operador Referencia" in df.columns
    assert df.iloc[0]["Operador Referencia"] == ""
    assert df.iloc[0]["Plataforma"] == "Netrefer"
