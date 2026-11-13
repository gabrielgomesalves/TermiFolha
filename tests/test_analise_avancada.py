"""Correlação, outliers, normalização, faixas e gravação de gráficos em PNG."""
import math

import pandas as pd
import pytest


def _folha(app, **colunas):
    app.m.table = pd.DataFrame(colunas)
    app.m.col_types = {c: (int if str(app.m.table[c].dtype).startswith("int") else float) for c in colunas}
    app.m.apply_types()


# ---------------------------------------------------------------- correlação


def test_correlacao_perfeita_positiva_e_negativa(app):
    _folha(app, X=[1, 2, 3, 4, 5], Y=[2, 4, 6, 8, 10], Z=[5, 4, 3, 2, 1])

    assert app.m.correlation("X", "Y") == (pytest.approx(1.0), 5)
    assert app.m.correlation("X", "Z")[0] == pytest.approx(-1.0)


def test_spearman_apanha_relacao_monotona_nao_linear(app):
    _folha(app, X=[1, 2, 3, 4, 5], Q=[1, 4, 9, 16, 25])

    assert app.m.correlation("X", "Q", "pearson")[0] < 1.0
    assert app.m.correlation("X", "Q", "spearman")[0] == pytest.approx(1.0)


def test_correlacao_ignora_linhas_com_valores_em_falta(app):
    _folha(app, X=[1.0, 2.0, 3.0, float("nan"), 5.0], Y=[2.0, 4.0, 6.0, 8.0, float("nan")])

    r, n = app.m.correlation("X", "Y")

    assert n == 3 and r == pytest.approx(1.0)


def test_correlacao_metodo_desconhecido(app):
    _folha(app, X=[1, 2, 3], Y=[3, 1, 2])

    with pytest.raises(ValueError, match="desconhecido"):
        app.m.correlation("X", "Y", "kendall")


def test_spearman_com_empates_usa_ordens_medias(app):
    _folha(app, X=[1, 2, 2, 3], Y=[1, 2, 3, 4])

    assert app.m.correlation("X", "Y", "spearman")[0] == pytest.approx(0.9487, abs=1e-3)


def test_correlacao_com_poucos_pares_ou_coluna_constante_levanta(app):
    _folha(app, X=[1, 2], Y=[3, 4], K=[7, 7])
    with pytest.raises(ValueError, match="pelo menos 3"):
        app.m.correlation("X", "Y")

    _folha(app, X=[1, 2, 3], K=[7, 7, 7])
    with pytest.raises(ValueError, match="não varia"):
        app.m.correlation("X", "K")


@pytest.mark.parametrize("r, texto", [
    (0.95, "muito forte positiva"), (-0.85, "muito forte negativa"), (0.65, "forte positiva"),
    (-0.45, "moderada negativa"), (0.25, "fraca positiva"), (0.05, "muito fraca ou nula"), (float("nan"), "indefinida"),
])
def test_descricao_da_intensidade(app, r, texto):
    assert app.m.correlation_strength(r) == texto


def test_correlacao_pela_interface_com_duas_e_com_varias_colunas(app, capsys):
    app.dummy()

    app.run(app.m.stat_correlation, "Idade,Salario")
    detalhe = capsys.readouterr().out
    assert "Pearson" in detalhe and "Spearman" in detalhe and "7 pares" in detalhe

    app.m.table["Extra"] = app.m.table["Salario"] * 2
    app.m.col_types["Extra"] = float
    app.run(app.m.stat_correlation, "Idade,Salario,Extra")
    assert "Matriz de correlação" in capsys.readouterr().out


def test_correlacao_pela_interface_mostra_erro_sem_rebentar(app, capsys):
    _folha(app, X=[1, 2, 3], K=[7, 7, 7])

    app.run(app.m.stat_correlation, "X,K")

    assert "não varia" in capsys.readouterr().out


# ---------------------------------------------------------------- outliers


def test_outliers_pelo_iqr(app):
    _folha(app, N=[1, 2, 3, 4, 100])

    mascara, baixo, alto = app.m.detect_outliers("N")

    assert mascara.tolist() == [False, False, False, False, True] and (baixo, alto) == (-1.0, 7.0)


def test_outliers_pelo_desvio_padrao_e_factor(app):
    _folha(app, N=[1, 2, 3, 4, 100])

    assert not app.m.detect_outliers("N", "zscore")[0].any()  # com 5 pontos, 100 fica a ~1.8 desvios
    assert app.m.detect_outliers("N", "zscore", factor=1.5)[0].tolist()[-1] is True


def test_em_falta_nunca_e_anomalo_e_coluna_vazia_nao_tem_anomalos(app):
    _folha(app, N=[1.0, 2.0, 3.0, 4.0, float("nan"), 100.0], V=[float("nan")] * 6)

    assert app.m.detect_outliers("N")[0].tolist() == [False] * 5 + [True]
    assert not app.m.detect_outliers("V")[0].any()


def test_outliers_metodo_desconhecido(app):
    _folha(app, N=[1, 2, 3])

    with pytest.raises(ValueError, match="desconhecido"):
        app.m.detect_outliers("N", "magia")


def test_outliers_remover_linhas_pela_interface(app):
    _folha(app, N=[1, 2, 3, 4, 100])

    app.run(app.m.outliers, "N", "", "", "1")

    assert app.m.table["N"].tolist() == [1, 2, 3, 4] and app.m.table.index.tolist() == [0, 1, 2, 3]


def test_outliers_deixar_em_branco_pela_interface(app):
    _folha(app, N=[1, 2, 3, 4, 100])

    app.run(app.m.outliers, "N", "1", "1.5", "2")

    assert len(app.m.table) == 5 and pd.isna(app.m.table.loc[4, "N"]) and str(app.m.table["N"].dtype) == "Int64"


def test_outliers_so_ver_nao_altera_nem_regista_passo(app):
    _folha(app, N=[1, 2, 3, 4, 100])
    passos = len(app.m.undo_stack)

    app.run(app.m.outliers, "N", "", "", "")

    assert len(app.m.table) == 5 and len(app.m.undo_stack) == passos


def test_outliers_sem_anomalos_e_entradas_invalidas(app, capsys):
    _folha(app, N=[1, 2, 3, 4, 5])
    app.run(app.m.outliers, "N", "", "")
    assert "Nenhum valor anómalo" in capsys.readouterr().out

    app.run(app.m.outliers, "N", "9")  # método inválido
    app.run(app.m.outliers, "N", "1", "abc")  # factor inválido
    app.run(app.m.outliers, "N", "1", "0")  # factor abaixo do mínimo
    assert len(app.m.table) == 5


def test_outliers_em_branco_numa_coluna_calculada_e_recusado(app):
    _folha(app, N=[1, 2, 3, 4, 100])
    app.run(app.m.add_formula_column, "M", "N * 1")

    app.run(app.m.outliers, "M", "", "", "2")

    assert app.m.table["M"].notna().all()


# ---------------------------------------------------------------- normalizar


def test_normalizar_minmax_e_zscore(app):
    serie = pd.Series([10, 20, 30, float("nan")])

    minmax = app.m.normalize_series(serie, "minmax")
    assert minmax.tolist()[:3] == [0.0, 0.5, 1.0] and math.isnan(minmax[3])

    z = app.m.normalize_series(serie, "zscore")
    assert z.dropna().mean() == pytest.approx(0.0) and z.dropna().std() == pytest.approx(1.0)


@pytest.mark.parametrize("metodo", ["minmax", "zscore"])
def test_normalizar_coluna_constante_ou_vazia_levanta(app, metodo):
    with pytest.raises(ValueError, match="não varia"):
        app.m.normalize_series(pd.Series([5, 5, 5]), metodo)
    with pytest.raises(ValueError, match="não varia"):
        app.m.normalize_series(pd.Series([float("nan")] * 3), metodo)
    with pytest.raises(ValueError, match="desconhecido"):
        app.m.normalize_series(pd.Series([1, 2]), "x")


def test_normalizar_pela_interface_cria_coluna_sem_mexer_na_original(app):
    app.dummy()

    app.run(app.m.normalize_column, "Salario", "1", "")

    assert app.m.col_types["Salario_norm"] is float
    assert app.m.table["Salario_norm"].max() == 1.0 and app.m.table["Salario_norm"].min() == 0.0
    assert app.m.table.loc[0, "Salario"] == 1500.50


def test_normalizar_nomes_repetidos_recebem_sufixo_ou_sao_recusados(app):
    app.dummy()
    app.run(app.m.normalize_column, "Salario", "1", "")
    app.run(app.m.normalize_column, "Salario", "1", "")
    assert "Salario_norm_2" in app.m.table.columns

    app.run(app.m.normalize_column, "Salario", "2", "Nome")  # nome já existe
    assert "Nome" in app.m.table.columns and app.m.table["Nome"].tolist()[0] == "João"
    app.run(app.m.normalize_column, "Salario", "9")  # método inválido


def test_normalizar_coluna_constante_pela_interface(app, capsys):
    _folha(app, K=[3, 3, 3])

    app.run(app.m.normalize_column, "K", "1")

    assert "não varia" in capsys.readouterr().out and list(app.m.table.columns) == ["K"]


# ---------------------------------------------------------------- faixas


def test_faixas_com_limites_escolhidos(app):
    idades = pd.Series([0, 17, 18, 19, 30, 31, 65, 70, float("nan")])

    rotulos, limites = app.m.make_bins(idades, "edges", ["0", "18", "30", "65"])

    assert limites == [0.0, 18.0, 30.0, 65.0]
    assert rotulos.tolist()[:7] == ["0 a 18", "0 a 18", "0 a 18", "18 a 30", "18 a 30", "30 a 65", "30 a 65"]
    assert pd.isna(rotulos[7]) and pd.isna(rotulos[8])  # 70 fora dos limites; em falta continua em falta


def test_faixas_de_igual_largura_e_por_quantis(app):
    serie = pd.Series(range(1, 101), dtype="float64")

    largura, limites = app.m.make_bins(serie, "equal", 4)
    assert len(limites) == 5 and largura.value_counts().nunique() == 1  # 4 faixas de 25

    quantis, _ = app.m.make_bins(serie, "quantile", 4)
    assert quantis.value_counts().min() >= 24


@pytest.mark.parametrize("modo, spec, mensagem", [
    ("edges", ["5"], "pelo menos 2"),
    ("edges", ["5", "3"], "crescente"),
    ("edges", ["1", "1"], "crescente"),
    ("edges", ["a", "b"], "números"),
    ("equal", 1, "entre 2"),
    ("equal", 51, "entre 2"),
    ("quantile", "3", "entre 2"),
    ("magia", 3, "desconhecido"),
])
def test_faixas_parametros_invalidos(app, modo, spec, mensagem):
    with pytest.raises(ValueError, match=mensagem):
        app.m.make_bins(pd.Series([1.0, 2.0, 3.0]), modo, spec)


def test_faixas_coluna_constante_ou_quase_vazia(app):
    with pytest.raises(ValueError, match="não varia"):
        app.m.make_bins(pd.Series([4.0, 4.0, 4.0]), "equal", 3)
    with pytest.raises(ValueError, match="suficientes"):
        app.m.make_bins(pd.Series([1.0, float("nan")]), "equal", 3)


def test_faixas_pela_interface_com_limites(app, capsys):
    app.dummy()

    app.run(app.m.bin_column, "Idade", "1", "0, 30, 60", "")

    assert app.m.table["Idade_faixa"].tolist()[:3] == ["0 a 30", "0 a 30", "30 a 60"]  # (a, b]: 30 fica na primeira
    assert pd.isna(app.m.table.loc[3, "Idade_faixa"]) and app.m.col_types["Idade_faixa"] is str
    assert "2 faixa(s)" in capsys.readouterr().out


def test_faixas_pela_interface_avisa_valores_fora_dos_limites(app, capsys):
    app.dummy()

    app.run(app.m.bin_column, "Idade", "1", "0,30", "Grupo")

    saida = capsys.readouterr().out
    assert "ficaram fora dos limites" in saida and "Grupo" in app.m.table.columns


@pytest.mark.parametrize("respostas", [
    ("Idade", "1", "30,10"),  # limites por ordem errada
    ("Idade", "2", "1"),  # nº de faixas fora do intervalo
    ("Idade", "2", "abc"),
    ("Idade", "9"),
    ("Idade", "1", "0,30", "Nome"),  # nome já existe
])
def test_faixas_pela_interface_entradas_invalidas_nao_criam_coluna(app, respostas):
    app.dummy()
    colunas = list(app.m.table.columns)

    app.run(app.m.bin_column, *respostas)

    assert list(app.m.table.columns) == colunas


def test_faixas_pela_interface_iguais_e_quantis(app):
    app.dummy()

    app.run(app.m.bin_column, "Salario", "2", "3", "")
    app.run(app.m.bin_column, "Salario", "3", "2", "")

    assert {"Salario_faixa", "Salario_faixa_2"} <= set(app.m.table.columns)


# ---------------------------------------------------------------- gráficos em PNG

_GRAFICOS = [
    ("scatter", ["Idade,Salario"]), ("area", ["Salario"]), ("histogram", ["Salario", ""]), ("boxplot", ["Salario"]),
    ("correlation_heatmap", ["Idade,Salario"]), ("moving_average", ["Salario", ""]), ("polynomial_trend", ["Salario", ""]),
    ("bars", ["1", "Salario"]), ("lines", ["1", "Salario"]), ("pie", ["Cidade", "Salario", "1"]),
]


@pytest.mark.parametrize("nome, respostas", _GRAFICOS)
def test_todos_os_graficos_gravam_png_quando_activado(app, tmp_path, monkeypatch, nome, respostas):
    app.dummy()
    pasta = tmp_path / "png"
    pasta.mkdir()
    monkeypatch.setattr(app.m, "plot_png_dir", str(pasta))

    app.run(getattr(app.m, f"plot_{nome}"), *respostas)

    ficheiros = list(pasta.glob("*.png"))
    assert len(ficheiros) == 1 and ficheiros[0].name.startswith(nome) and ficheiros[0].stat().st_size > 1000


def test_sem_png_o_grafico_abre_em_janela(app, monkeypatch):
    app.dummy()
    chamadas = []
    monkeypatch.setattr(app.m.plt, "show", lambda: chamadas.append(1))

    app.run(app.m.plot_area, "Salario")

    assert chamadas == [1]


def test_dois_graficos_seguidos_nao_se_sobrepoem(app, tmp_path, monkeypatch):
    app.dummy()
    monkeypatch.setattr(app.m, "plot_png_dir", str(tmp_path))
    monkeypatch.setattr(app.m.time, "strftime", lambda _fmt: "FIXO")

    app.run(app.m.plot_area, "Salario")
    app.run(app.m.plot_area, "Salario")

    assert sorted(p.name for p in tmp_path.glob("*.png")) == ["area_FIXO.png", "area_FIXO_2.png"]


def test_falha_a_gravar_png_avisa_sem_rebentar(app, tmp_path, monkeypatch, capsys):
    app.dummy()
    monkeypatch.setattr(app.m, "plot_png_dir", str(tmp_path / "nao_existe"))

    app.run(app.m.plot_area, "Salario")

    assert "Erro ao gravar" in capsys.readouterr().out


def test_ligar_e_desligar_gravacao_em_png(app, tmp_path, monkeypatch):
    monkeypatch.setattr(app.m, "plot_png_dir", None)
    monkeypatch.chdir(tmp_path)

    app.run(app.m.toggle_png_export, "saida")
    assert app.m.plot_png_dir == "saida" and (tmp_path / "saida").is_dir()

    app.run(app.m.toggle_png_export)
    assert app.m.plot_png_dir is None


def test_ligar_png_numa_pasta_impossivel_nao_liga(app, tmp_path, monkeypatch):
    monkeypatch.setattr(app.m, "plot_png_dir", None)
    ficheiro = tmp_path / "ficheiro"
    ficheiro.write_text("x")

    app.run(app.m.toggle_png_export, str(ficheiro / "sub"))

    assert app.m.plot_png_dir is None
