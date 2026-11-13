"""Regras de validação, conversão de tipo de coluna e procurar/substituir."""
import pandas as pd
import pytest

# ---------------------------------------------------------------- regras de validação


def test_regra_de_minimo_e_maximo(app):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"min": "18", "max": "65"}

    assert app.m.rule_violation("Idade", 17) == "abaixo do mínimo (18)"
    assert app.m.rule_violation("Idade", 66) == "acima do máximo (65)"
    assert app.m.rule_violation("Idade", 30) is None
    assert app.m.rule_violation("Idade", 18) is None and app.m.rule_violation("Idade", 65) is None  # limites inclusivos
    assert app.m.rule_violation("Idade", pd.NA) is None  # em branco só falha se for obrigatório
    assert app.m.rule_violation("Nome", "qualquer coisa") is None  # sem regra


def test_regra_obrigatorio(app):
    app.dummy()
    app.m.meta["rules"]["Nome"] = {"required": True}

    for vazio in ("", "  ", None, float("nan"), pd.NA):
        assert app.m.rule_violation("Nome", vazio) == "o valor é obrigatório"
    assert app.m.rule_violation("Nome", "Ana") is None


def test_regra_valores_permitidos_texto_numero_e_data(app):
    app.dummy()
    app.m.meta["rules"] = {
        "Cidade": {"allowed": ["Lisboa", "Porto"]},
        "Idade": {"allowed": ["25", "30"]},
        "Admissao": {"allowed": ["15/01/2020"]},
    }

    assert app.m.rule_violation("Cidade", "porto") is None  # sem distinguir maiúsculas
    assert "fora dos permitidos" in app.m.rule_violation("Cidade", "Faro")
    assert app.m.rule_violation("Idade", 25.0) is None
    assert app.m.rule_violation("Idade", 26) is not None
    assert app.m.rule_violation("Admissao", pd.Timestamp("2020-01-15")) is None
    assert app.m.rule_violation("Admissao", pd.Timestamp("2021-01-15")) is not None


def test_regra_de_datas(app):
    app.dummy()
    app.m.meta["rules"]["Admissao"] = {"min": "01/01/2019", "max": "31/12/2021"}

    assert app.m.rule_violation("Admissao", pd.Timestamp("2018-12-31")) is not None
    assert app.m.rule_violation("Admissao", pd.Timestamp("2022-01-01")) is not None
    assert app.m.rule_violation("Admissao", pd.Timestamp("2020-06-01")) is None


def test_limite_corrompido_no_estado_e_ignorado(app):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"min": "abc"}

    assert app.m.rule_violation("Idade", 5) is None


def test_definir_regra_pela_interface(app):
    app.dummy()

    app.run(app.m.set_rule, "Idade", "18", "65", "", "n")

    assert app.m.meta["rules"]["Idade"] == {"min": "18", "max": "65"}


def test_definir_regra_texto_com_permitidos_e_obrigatorio(app):
    app.dummy()

    app.run(app.m.set_rule, "Cidade", "Lisboa, Porto", "s")

    assert app.m.meta["rules"]["Cidade"] == {"allowed": ["Lisboa", "Porto"], "required": True}


@pytest.mark.parametrize("respostas", [
    ("Idade", "abc"),  # mínimo inválido
    ("Idade", "10", "xyz"),  # máximo inválido
    ("Idade", "50", "10"),  # mínimo > máximo
    ("Idade", "", "", "abc", "n"),  # permitido que não é inteiro
    ("Idade", "", "", "", "n"),  # nada definido
    ("Nome", "", "n"),  # nada definido (texto)
])
def test_regra_invalida_ou_vazia_nao_e_guardada(app, respostas):
    app.dummy()

    app.run(app.m.set_rule, *respostas)

    assert app.m.meta["rules"] == {}


def test_adicionar_linha_recusa_valor_que_viola_a_regra(app, capsys):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"min": "18"}

    app.run(app.m.add_row, "Zé", "10", "20", "1.0", "Faro", "01/01/2024", "sair")

    assert "Regra da coluna 'Idade'" in capsys.readouterr().out
    assert len(app.m.table) == 9 and app.m.table.loc[8, "Idade"] == 20  # 10 recusado, 20 aceite


def test_editar_celula_recusa_valor_que_viola_a_regra(app, capsys):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"max": "50"}

    app.run(app.m.edit_cell, "0", "Idade", "99", "40")

    assert "Regra da coluna" in capsys.readouterr().out
    assert app.m.table.loc[0, "Idade"] == 40


def test_obrigatorio_recusa_texto_em_branco_ao_adicionar(app):
    app.dummy()
    app.m.meta["rules"]["Nome"] = {"required": True}

    app.run(app.m.add_row, "", "Zé", "20", "1.0", "Faro", "01/01/2024", "sair")

    assert app.m.table.loc[8, "Nome"] == "Zé"


def test_verificar_folha_lista_violacoes(app, capsys):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"min": "30"}

    violacoes = app.m.find_violations()
    app.m.check_rules()

    assert [(i, c) for i, c, _v, _m in violacoes] == [(0, "Idade"), (5, "Idade"), (6, "Idade")]  # a célula em falta não conta
    assert "3 violação" in capsys.readouterr().out


def test_verificar_folha_sem_violacoes(app, capsys):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"min": "1"}

    app.m.check_rules()

    assert "cumprem" in capsys.readouterr().out


def test_remover_regras_pela_interface(app):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"min": "1"}

    app.run(app.m.remove_rule, "Idade")

    assert app.m.meta["rules"] == {}


def test_menu_de_regras_navega(app, capsys):
    app.dummy()
    app.m.meta["rules"]["Idade"] = {"min": "1"}

    app.run(app.m.menu_rules, "3", "0")

    assert "Idade: min=1" in capsys.readouterr().out


# ---------------------------------------------------------------- conversão de tipo


def _serie(*valores):
    return pd.Series(list(valores), dtype=object)


def test_converter_texto_em_numeros(app):
    convertida, falhou = app.m.convert_series(_serie("1", " 2 ", "x", None), str, int)
    assert convertida.tolist()[:2] == [1, 2] and pd.isna(convertida[2]) and str(convertida.dtype) == "Int64"
    assert falhou.tolist() == [False, False, True, False]  # em branco não é falha

    convertida, falhou = app.m.convert_series(_serie("1,5", "2.25", "abc"), str, float)
    assert convertida.tolist()[:2] == [1.5, 2.25] and falhou.tolist() == [False, False, True]


def test_decimal_para_inteiro_so_se_for_inteiro(app):
    convertida, falhou = app.m.convert_series(pd.Series([2.0, 2.5, None]), float, int)

    assert convertida[0] == 2 and pd.isna(convertida[1]) and falhou.tolist() == [False, True, False]


def test_converter_para_texto_e_datas(app):
    texto, _ = app.m.convert_series(pd.Series([25, None], dtype="Int64"), int, str)
    assert texto[0] == "25" and pd.isna(texto[1])

    texto, _ = app.m.convert_series(pd.Series([1500.5, 2000.0]), float, str)
    assert texto.tolist() == ["1500.5", "2000"]

    texto, _ = app.m.convert_series(pd.Series(pd.to_datetime(["2020-01-15"])), app.m.DATE_TYPE, str)
    assert texto[0] == "15/01/2020"

    datas, falhou = app.m.convert_series(_serie("15/01/2020", "2020-01-15", "x"), str, app.m.DATE_TYPE)
    assert datas[0] == pd.Timestamp("2020-01-15") and falhou.tolist() == [False, True, True]


def test_inteiro_para_decimal_e_sem_perdas(app):
    convertida, falhou = app.m.convert_series(pd.Series([1, None], dtype="Int64"), int, float)

    assert convertida.tolist()[0] == 1.0 and not falhou.any() and str(convertida.dtype) == "float64"


@pytest.mark.parametrize("origem, destino", [(int, "date"), ("date", float)])
def test_numeros_e_datas_nao_se_convertem_directamente(app, origem, destino):
    with pytest.raises(ValueError, match="números e datas"):
        app.m.convert_series(pd.Series([1]), origem, destino)


def _folha_texto(app):
    app.m.table = pd.DataFrame({"N": ["1", "2", "x"], "T": ["a", "b", "c"]})
    app.m.col_types = {"N": str, "T": str}


def test_converter_coluna_pela_interface_com_falhas_confirmadas(app):
    _folha_texto(app)
    app.m.meta["rules"]["N"] = {"required": True}

    app.run(app.m.convert_column_type, "N", "2", "s")

    assert app.m.col_types["N"] is int and app.m.table["N"].tolist()[:2] == [1, 2] and pd.isna(app.m.table.loc[2, "N"])
    assert "N" not in app.m.meta["rules"]  # regras do tipo antigo deixam de se aplicar


def test_converter_coluna_cancelado_nao_altera(app):
    _folha_texto(app)

    app.run(app.m.convert_column_type, "N", "2", "n")

    assert app.m.col_types["N"] is str and app.m.table["N"].tolist() == ["1", "2", "x"]


def test_converter_sem_falhas_nao_pede_confirmacao(app):
    _folha_texto(app)
    app.m.table.loc[2, "N"] = "3"

    app.run(app.m.convert_column_type, "N", "2")

    assert app.m.table["N"].tolist() == [1, 2, 3]


def test_converter_para_o_mesmo_tipo_ou_coluna_calculada_e_recusado(app):
    app.dummy()
    app.run(app.m.convert_column_type, "Idade", "2")
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.convert_column_type, "Dobro")

    assert app.m.col_types["Idade"] is int and app.m.col_types["Dobro"] is int


def test_converter_numero_em_data_e_recusado(app):
    app.dummy()

    app.run(app.m.convert_column_type, "Idade", "4")

    assert app.m.col_types["Idade"] is int


def test_converter_pode_ser_desfeito(app):
    _folha_texto(app)
    app.m.commit()

    app.run(app.m.convert_column_type, "N", "2", "s")
    app.m.undo_last()

    assert app.m.col_types["N"] is str and app.m.table["N"].tolist() == ["1", "2", "x"]


# ---------------------------------------------------------------- procurar e substituir


def test_substituir_parte_do_texto_sem_distinguir_maiusculas(app):
    app.dummy()

    plano = app.m.plan_replace("lisboa", "Madrid", ["Cidade"])

    novo, mascara = plano["Cidade"]
    assert int(mascara.sum()) == 3 and novo.tolist()[0] == "Madrid" and novo.tolist()[1] == "Porto"
    assert app.m.table.loc[0, "Cidade"] == "Lisboa"  # o plano não altera a folha


def test_substituir_distinguindo_maiusculas_e_celula_inteira(app):
    app.dummy()

    assert app.m.plan_replace("lisboa", "X", ["Cidade"], case_sensitive=True) == {}
    assert app.m.plan_replace("Lis", "X", ["Cidade"], whole_cell=True) == {}  # parte não é célula inteira
    novo, _ = app.m.plan_replace("Lisboa", "X", ["Cidade"], whole_cell=True)["Cidade"]
    assert novo.tolist()[0] == "X"
    novo, _ = app.m.plan_replace("Lis", "X", ["Cidade"])["Cidade"]
    assert novo.tolist()[0] == "Xboa"


def test_o_texto_procurado_e_literal_e_a_substituicao_tambem(app):
    app.m.table = pd.DataFrame({"T": ["a.b", "axb", "C:\\dir"]})
    app.m.col_types = {"T": str}

    novo, mascara = app.m.plan_replace(".", "\\1", ["T"])["T"]

    assert mascara.tolist() == [True, False, False]  # '.' não é regex
    assert novo.tolist()[0] == "a\\1b"  # '\1' não é referência de grupo


def test_substituir_numeros_e_datas_so_em_celulas_iguais(app):
    app.dummy()

    novo, mascara = app.m.plan_replace("25", "26", ["Idade"])["Idade"]
    assert int(mascara.sum()) == 2 and novo[0] == 26

    novo, mascara = app.m.plan_replace("2000", "2100,5", ["Salario"])["Salario"]
    assert novo[1] == 2100.5  # aceita vírgula decimal

    novo, _ = app.m.plan_replace("15/01/2020", "16/01/2020", ["Admissao"])["Admissao"]
    assert novo[0] == pd.Timestamp("2020-01-16")

    assert app.m.plan_replace("abc", "1", ["Idade", "Salario", "Admissao"]) == {}


def test_substituicao_invalida_para_o_tipo_levanta(app):
    app.dummy()

    with pytest.raises(ValueError, match="não é válido"):
        app.m.plan_replace("25", "abc", ["Idade"])


def test_substituir_por_vazio_em_numerico_deixa_em_branco(app):
    app.dummy()

    novo, _ = app.m.plan_replace("25", "", ["Idade"])["Idade"]

    assert pd.isna(novo[0]) and pd.isna(novo[6])


def test_procurar_e_substituir_pela_interface_em_todas_as_colunas(app):
    app.dummy()

    app.run(app.m.find_replace, "", "João", "Joao", "n", "n", "s")

    assert app.m.table["Nome"].tolist()[0] == "Joao" and app.m.table["Nome"].tolist()[6] == "Joao"
    assert str(app.m.table["Idade"].dtype) == "Int64"


def test_substituir_cancelado_ou_sem_resultados_nao_altera(app):
    app.dummy()

    app.run(app.m.find_replace, "Nome", "João", "Joao", "n", "n", "n")
    app.run(app.m.find_replace, "Nome", "Zzz", "Joao", "n", "n")

    assert app.m.table["Nome"].tolist()[0] == "João"


def test_substituir_que_viola_regra_nao_altera_nada(app, capsys):
    app.dummy()
    app.m.meta["rules"]["Cidade"] = {"allowed": ["Lisboa", "Porto", "Faro", "Coimbra", "Braga"]}

    app.run(app.m.find_replace, "Cidade", "Lisboa", "Madrid", "s", "n")

    assert "violaria" in capsys.readouterr().out
    assert "Madrid" not in app.m.table["Cidade"].tolist()


def test_substituir_ignora_colunas_calculadas(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.find_replace, "", "50", "51", "s", "n", "s")  # em todas: só Idade (e nenhuma 50)
    assert app.m.table.loc[0, "Dobro"] == 50

    app.run(app.m.find_replace, "Dobro")  # directamente: recusado
    assert app.m.meta["formulas"] == {"Dobro": "Idade * 2"}


def test_substituir_pode_ser_desfeito(app):
    app.dummy()
    app.run(app.m.find_replace, "Nome", "João", "Joao", "n", "n", "s")

    app.m.undo_last()

    assert app.m.table["Nome"].tolist()[0] == "João"
