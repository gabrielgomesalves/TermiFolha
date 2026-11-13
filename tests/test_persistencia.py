import json
import os
import pickle
import subprocess
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent


def _caminho_por_omissao(tmp_path):
    codigo = f"import sys; sys.path.insert(0, {str(RAIZ)!r}); import TermiFolha as T; print(T.FICHEIRO_ESTADO); print(T.FICHEIRO_ESTADO_LEGADO)"
    saida = subprocess.run([sys.executable, "-c", codigo], cwd=tmp_path, capture_output=True, text=True, check=True).stdout.split("\n")
    return Path(saida[0]), Path(saida[1])


def test_ficheiros_de_estado_ficam_ao_lado_do_programa(tmp_path):
    json_, legado = _caminho_por_omissao(tmp_path)

    assert json_.parent == RAIZ and json_.name == ".termifolha_estado.json"
    assert legado.parent == RAIZ and legado.name == ".termifolha_estado.pkl"


def test_estado_e_json_e_nao_pickle(app):
    app.dummy()

    dados = json.loads(Path(app.m.FICHEIRO_ESTADO).read_text(encoding="utf-8"))

    assert dados["versao"] == 1
    assert dados["colunas"] == ["Nome", "Idade", "Salario", "Cidade", "Admissao"]
    assert dados["tipos"] == {"Nome": "str", "Idade": "int", "Salario": "float", "Cidade": "str", "Admissao": "date"}
    assert dados["dados"][3][1] is None  # valor em falta
    assert dados["dados"][0][4].startswith("2020-01-15")


def test_estado_guardado_e_recarregado_sem_perdas(app):
    app.dummy()
    original = app.m.tabela.copy()
    tipos_originais = dict(app.m.tipos)

    app.m.tabela = pd.DataFrame()
    app.m.tipos = {}
    app.m.carregar_estado()

    t = app.m.tabela
    assert list(t.columns) == list(original.columns)
    assert t["Nome"].tolist() == original["Nome"].tolist()
    assert str(t["Idade"].dtype) == "Int64"
    assert int(t["Idade"].isna().sum()) == 1
    assert t["Salario"].tolist() == original["Salario"].tolist()
    assert t["Admissao"].tolist() == original["Admissao"].tolist()
    assert app.m.tipos == tipos_originais
    assert len(app.m.undo) == 1


def test_folha_sem_linhas_e_recarregada(app):
    app.run(app.m.nova_folha, "A", "2", "B", "4", "")

    app.m.tabela = pd.DataFrame()
    app.m.tipos = {}
    app.m.carregar_estado()

    assert list(app.m.tabela.columns) == ["A", "B"]
    assert app.m.tabela.empty
    assert app.m.tipos == {"A": int, "B": app.m.TIPO_DATA}


def test_gravacao_nao_deixa_ficheiros_temporarios(app, tmp_path):
    app.dummy()

    assert sorted(f.name for f in tmp_path.iterdir()) == ["estado.json"]


def test_estado_corrompido_e_preservado_e_utilizador_e_avisado(app, capsys):
    caminho = Path(app.m.FICHEIRO_ESTADO)
    caminho.write_bytes(b"isto nao e json")

    app.m.carregar_estado()

    assert "corromp" in capsys.readouterr().out.lower()
    assert app.m.tabela.empty
    assert not caminho.exists()
    assert Path(str(caminho) + ".corrompido").read_bytes() == b"isto nao e json"


def test_json_valido_com_estrutura_errada_e_tratado_como_corrompido(app, capsys):
    caminho = Path(app.m.FICHEIRO_ESTADO)
    caminho.write_text(json.dumps({"versao": 1, "colunas": ["A", "B"], "tipos": {}, "dados": [[1]]}), encoding="utf-8")

    app.m.carregar_estado()

    assert "corromp" in capsys.readouterr().out.lower()
    assert app.m.tabela.empty


def test_versao_desconhecida_nao_e_carregada(app, capsys):
    caminho = Path(app.m.FICHEIRO_ESTADO)
    caminho.write_text(json.dumps({"versao": 99, "colunas": [], "tipos": {}, "dados": []}), encoding="utf-8")

    app.m.carregar_estado()

    assert "corromp" in capsys.readouterr().out.lower()


def test_falha_a_gravar_estado_avisa_o_utilizador(app, capsys, tmp_path):
    app.m.tabela = pd.DataFrame({"A": [1]})
    app.m.FICHEIRO_ESTADO = str(tmp_path / "nao_existe" / "estado.json")  # pasta inexistente

    assert app.m.guardar_estado() is False

    assert "não foi possível guardar" in capsys.readouterr().out.lower()


def test_migra_estado_antigo_em_pickle(app, capsys):
    app.dummy()
    esperado = app.m.tabela.copy()
    tipos = dict(app.m.tipos)
    os.remove(app.m.FICHEIRO_ESTADO)
    legado = Path(app.m.FICHEIRO_ESTADO_LEGADO)
    legado.write_bytes(pickle.dumps((esperado, tipos)))
    app.m.tabela = pd.DataFrame()
    app.m.tipos = {}

    app.m.carregar_estado()

    assert "migrado" in capsys.readouterr().out.lower()
    assert app.m.tabela["Nome"].tolist() == esperado["Nome"].tolist()
    assert app.m.tipos == tipos
    assert Path(app.m.FICHEIRO_ESTADO).exists()
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
        legado = Path(app.m.FICHEIRO_ESTADO_LEGADO)
        legado.write_bytes(pickle.dumps(classe(str(marcador))))

        app.m.carregar_estado()

        assert not marcador.exists(), f"{classe.__name__} foi executado!"
        assert "corromp" in capsys.readouterr().out.lower()
        assert app.m.tabela.empty
        assert not legado.exists()
        os.remove(str(legado) + ".corrompido")


def test_limpar_folha_apaga_o_estado_guardado(app):
    app.dummy()
    assert Path(app.m.FICHEIRO_ESTADO).exists()

    app.run(app.m.limpar_folha, "s")

    assert not Path(app.m.FICHEIRO_ESTADO).exists()
