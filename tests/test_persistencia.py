import json
import os
import pickle
import subprocess
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent


def _caminho_por_omissao(tmp_path):
    codigo = f"import sys; sys.path.insert(0, {str(RAIZ)!r}); import TermiFolha as T; print(T.STATE_FILE); print(T.LEGACY_STATE_FILE)"
    saida = subprocess.run([sys.executable, "-c", codigo], cwd=tmp_path, capture_output=True, text=True, check=True).stdout.split("\n")
    return Path(saida[0]), Path(saida[1])


def test_ficheiros_de_estado_ficam_ao_lado_do_programa(tmp_path):
    json_, legado = _caminho_por_omissao(tmp_path)

    assert json_.parent == RAIZ and json_.name == ".termifolha_estado.json"
    assert legado.parent == RAIZ and legado.name == ".termifolha_estado.pkl"


def test_estado_e_json_e_nao_pickle(app):
    app.dummy()

    data_rows = json.loads(Path(app.m.STATE_FILE).read_text(encoding="utf-8"))

    assert data_rows["versao"] == 1
    assert data_rows["colunas"] == ["Nome", "Idade", "Salario", "Cidade", "Admissao"]
    assert data_rows["tipos"] == {"Nome": "str", "Idade": "int", "Salario": "float", "Cidade": "str", "Admissao": "date"}
    assert data_rows["dados"][3][1] is None  # valor em falta
    assert data_rows["dados"][0][4].startswith("2020-01-15")


def test_estado_guardado_e_recarregado_sem_perdas(app):
    app.dummy()
    original = app.m.table.copy()
    tipos_originais = dict(app.m.col_types)

    app.m.table = pd.DataFrame()
    app.m.col_types = {}
    app.m.load_state()

    t = app.m.table
    assert list(t.columns) == list(original.columns)
    assert t["Nome"].tolist() == original["Nome"].tolist()
    assert str(t["Idade"].dtype) == "Int64"
    assert int(t["Idade"].isna().sum()) == 1
    assert t["Salario"].tolist() == original["Salario"].tolist()
    assert t["Admissao"].tolist() == original["Admissao"].tolist()
    assert app.m.col_types == tipos_originais
    assert len(app.m.undo_stack) == 1


def test_folha_sem_linhas_e_recarregada(app):
    app.run(app.m.new_sheet, "A", "2", "B", "4", "")

    app.m.table = pd.DataFrame()
    app.m.col_types = {}
    app.m.load_state()

    assert list(app.m.table.columns) == ["A", "B"]
    assert app.m.table.empty
    assert app.m.col_types == {"A": int, "B": app.m.DATE_TYPE}


def test_gravacao_nao_deixa_ficheiros_temporarios(app, tmp_path):
    app.dummy()

    assert sorted(f.name for f in tmp_path.iterdir()) == ["estado.json"]


def test_estado_corrompido_e_preservado_e_utilizador_e_avisado(app, capsys):
    path = Path(app.m.STATE_FILE)
    path.write_bytes(b"isto nao e json")

    app.m.load_state()

    assert "corromp" in capsys.readouterr().out.lower()
    assert app.m.table.empty
    assert not path.exists()
    assert Path(str(path) + ".corrompido").read_bytes() == b"isto nao e json"


def test_json_valido_com_estrutura_errada_e_tratado_como_corrompido(app, capsys):
    path = Path(app.m.STATE_FILE)
    path.write_text(json.dumps({"versao": 1, "colunas": ["A", "B"], "tipos": {}, "dados": [[1]]}), encoding="utf-8")

    app.m.load_state()

    assert "corromp" in capsys.readouterr().out.lower()
    assert app.m.table.empty


def test_versao_desconhecida_nao_e_carregada(app, capsys):
    path = Path(app.m.STATE_FILE)
    path.write_text(json.dumps({"versao": 99, "colunas": [], "tipos": {}, "dados": []}), encoding="utf-8")

    app.m.load_state()

    assert "corromp" in capsys.readouterr().out.lower()


def test_falha_a_gravar_estado_avisa_o_utilizador(app, capsys, tmp_path):
    app.m.table = pd.DataFrame({"A": [1]})
    app.m.STATE_FILE = str(tmp_path / "nao_existe" / "estado.json")  # pasta inexistente

    assert app.m.save_state() is False

    assert "não foi possível guardar" in capsys.readouterr().out.lower()


def test_migra_estado_antigo_em_pickle(app, capsys):
    app.dummy()
    esperado = app.m.table.copy()
    col_types = dict(app.m.col_types)
    os.remove(app.m.STATE_FILE)
    legado = Path(app.m.LEGACY_STATE_FILE)
    legado.write_bytes(pickle.dumps((esperado, col_types)))
    app.m.table = pd.DataFrame()
    app.m.col_types = {}

    app.m.load_state()

    assert "migrado" in capsys.readouterr().out.lower()
    assert app.m.table["Nome"].tolist() == esperado["Nome"].tolist()
    assert app.m.col_types == col_types
    assert Path(app.m.STATE_FILE).exists()
    assert not legado.exists()
    assert Path(str(legado) + ".migrado").exists()


class _Malicioso:
    def __init__(self, marcador):
        self.marcador = marcador

    def __reduce__(self):
        return (os.system, (f'echo pwned > "{self.marcador}"',))


class _MaliciosoPandasEval:
    def __init__(self, marcador):
        self.marcador = marcador

    def __reduce__(self):
        return (pd.eval, (f"__import__('pathlib').Path(r'{self.marcador}').write_text('x')",))


def test_pickle_malicioso_nao_e_executado(app, tmp_path, capsys):
    for i, classe in enumerate((_Malicioso, _MaliciosoPandasEval)):
        marcador = tmp_path / f"pwned{i}.txt"
        legado = Path(app.m.LEGACY_STATE_FILE)
        legado.write_bytes(pickle.dumps(classe(str(marcador))))

        app.m.load_state()

        assert not marcador.exists(), f"{classe.__name__} foi executado!"
        assert "corromp" in capsys.readouterr().out.lower()
        assert app.m.table.empty
        assert not legado.exists()
        os.remove(str(legado) + ".corrompido")


def test_limpar_folha_apaga_o_estado_guardado(app):
    app.dummy()
    assert Path(app.m.STATE_FILE).exists()

    app.run(app.m.clear_sheet, "s")

    assert not Path(app.m.STATE_FILE).exists()
