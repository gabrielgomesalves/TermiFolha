"""Colunas calculadas: avaliador seguro, recalculo automático e protecções."""
import pandas as pd
import pytest


def _avaliar(app, expressao):
    return app.m.evaluate_formula(expressao, app.m.table, app.m.col_types)


def test_aritmetica_com_colunas_e_tipo_do_resultado(app):
    app.dummy()

    serie, tipo = _avaliar(app, "Salario * 2")
    assert tipo is float and serie[0] == 3001.0

    serie, tipo = _avaliar(app, "Idade * 2 + 1")
    assert tipo is int and serie[0] == 51 and pd.isna(serie[3])  # em falta propaga-se

    assert _avaliar(app, "Idade / 5")[1] is float
    assert _avaliar(app, "sqrt(Idade)")[1] is float
    assert _avaliar(app, "Idade // 7")[1] is int


def test_nomes_sem_distinguir_maiusculas_e_com_parentesis_rectos(app):
    app.dummy()
    app.m.table = app.m.table.rename(columns={"Salario": "Preço unitário"})
    app.m.col_types["Preço unitário"] = app.m.col_types.pop("Salario")

    serie, _ = _avaliar(app, "[preço unitário] + idade")

    assert serie[0] == 1525.5


def test_funcoes_e_comparacoes(app):
    app.dummy()

    assert _avaliar(app, "se(Idade >= 30, 1, 0)")[0].tolist() == [0, 1, 1, 0, 1, 0, 0, 1]  # em falta -> falso
    assert _avaliar(app, "round(Salario / 7, 1)")[0][0] == 214.4
    assert _avaliar(app, "max(Idade, 30)")[0][0] == 30
    assert _avaliar(app, "abs(0 - Idade)")[0][0] == 25
    assert _avaliar(app, "Idade > 26 and Salario < 2100")[0].tolist() == [0, 1, 1, 0, 0, 1, 0, 1]


def test_texto_e_datas(app):
    app.dummy()

    texto, tipo = _avaliar(app, 'Nome + " (" + Cidade + ")"')
    assert tipo is str and texto[1] == "Maria (Porto)"

    ano, tipo = _avaliar(app, "year(Admissao)")
    assert tipo is int and ano[0] == 2020

    dias, _ = _avaliar(app, "days(Admissao, Admissao)")
    assert dias.eq(0).all()


def test_divisao_por_zero_e_nan(app):
    app.dummy()

    serie, tipo = _avaliar(app, "Idade / 0")

    assert tipo is float and serie.isna().all()  # infinito vira em falta


@pytest.mark.parametrize("expressao, parte_da_mensagem", [
    ("", "vazia"),
    ("Salario *", "sintaxe"),
    ("Inexistente + 1", "desconhecida"),
    ("[Inexistente] + 1", "inexistente"),
    ("foo(Idade)", "desconhecida"),
    ("abs(Idade, 2)", "argumento"),
    ("round(x=Idade)", "chamada inválida"),
    ("1 < Idade < 5", "encadeadas"),
    ("True", "números e texto"),
    ("None", "números e texto"),
])
def test_formulas_invalidas_dao_erro_claro(app, expressao, parte_da_mensagem):
    app.dummy()

    with pytest.raises(app.m.FormulaError, match=parte_da_mensagem):
        _avaliar(app, expressao)


@pytest.mark.parametrize("expressao", [
    "__import__('os').system('echo pwned')",
    "Nome.upper()",
    "(1).__class__",
    "lambda: 1",
    "[x for x in Idade]",
    "Idade[0]",
    "{1: 2}",
    "Idade if Idade else 0",
    "open('x')",
    "eval('1')",
    "abs.__call__(1)",
    "f'{Idade}'",
    "Idade := 3",
])
def test_nunca_executa_codigo_arbitrario(app, expressao, tmp_path):
    app.dummy()

    with pytest.raises(app.m.FormulaError):
        _avaliar(app, expressao)

    assert not list(tmp_path.glob("pwned*"))


def test_limites_de_tamanho_e_complexidade(app):
    app.dummy()

    with pytest.raises(app.m.FormulaError, match="caracteres"):
        _avaliar(app, "1+" * 300 + "1")
    with pytest.raises(app.m.FormulaError, match="complexa"):
        _avaliar(app, "+".join(["1"] * 150))


def test_potencias_gigantes_nao_bloqueiam(app):
    app.dummy()

    serie, tipo = _avaliar(app, "9 ** 9 ** 9 ** 9")

    assert tipo is float and serie.isna().all()  # overflow -> infinito -> em falta


def test_criar_coluna_calculada_pela_interface(app):
    app.dummy()

    app.run(app.m.add_formula_column, "Dobro", "Inexistente * 2", "Idade * 2")

    assert app.m.meta["formulas"] == {"Dobro": "Idade * 2"}
    assert app.m.col_types["Dobro"] is int and app.m.table.loc[0, "Dobro"] == 50


def test_nome_repetido_ou_cancelar_nao_cria_coluna(app):
    app.dummy()

    app.run(app.m.add_formula_column, "Nome")
    app.run(app.m.add_formula_column, "Novo", "sair")

    assert app.m.meta["formulas"] == {} and "Novo" not in app.m.table.columns


def test_recalculo_automatico_ao_editar_adicionar_e_remover_linhas(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.edit_cell, "0", "Idade", "40")
    assert app.m.table.loc[0, "Dobro"] == 80

    app.run(app.m.add_row, "Zé", "10", "100.0", "Faro", "01/01/2024", "sair")
    assert app.m.table.loc[8, "Dobro"] == 20  # a coluna calculada não é pedida

    app.run(app.m.remove_row, "0", "s")
    assert len(app.m.table) == 8 and app.m.table["Dobro"].notna().sum() == 7


def test_desfazer_repoe_valores_calculados(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")
    app.run(app.m.edit_cell, "0", "Idade", "40")

    app.m.undo_last()

    assert app.m.table.loc[0, "Dobro"] == 50 and app.m.table.loc[0, "Idade"] == 25


def test_nao_deixa_editar_coluna_calculada(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.edit_cell, "0", "Dobro")  # recusa antes de pedir o valor

    assert app.m.table.loc[0, "Dobro"] == 50


def test_nao_deixa_renomear_nem_remover_coluna_usada_por_formula(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.rename_column, "Idade")
    app.run(app.m.remove_column, "Idade")

    assert "Idade" in app.m.table.columns and "Dobro" in app.m.meta["formulas"]


def test_renomear_coluna_calculada_move_a_formula(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.rename_column, "Dobro", "Duplo")

    assert app.m.meta["formulas"] == {"Duplo": "Idade * 2"} and "Duplo" in app.m.table.columns


def test_remover_coluna_calculada(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.remove_column, "Dobro", "s")

    assert app.m.meta["formulas"] == {} and "Dobro" not in app.m.table.columns


def test_formula_que_usa_outra_formula(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")
    app.run(app.m.add_formula_column, "Quadruplo", "Dobro * 2")
    app.run(app.m.edit_cell, "0", "Idade", "10")

    assert app.m.table.loc[0, "Quadruplo"] == 40
    assert app.m.formula_users("Dobro") == ["Quadruplo"]

    app.run(app.m.freeze_formula, "Dobro")  # usada por outra: recusa
    assert "Dobro" in app.m.meta["formulas"]


def test_converter_formula_em_valores_permite_editar(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.run(app.m.freeze_formula, "Dobro", "s")
    app.run(app.m.edit_cell, "0", "Idade", "40")

    assert app.m.meta["formulas"] == {}
    assert app.m.table.loc[0, "Dobro"] == 50  # já não recalcula


def test_formula_guardada_e_recarregada(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")

    app.m.reset_workbook()
    app.m.load_state()

    assert app.m.meta["formulas"] == {"Dobro": "Idade * 2"} and app.m.col_types["Dobro"] is int
    app.run(app.m.edit_cell, "1", "Idade", "50")
    assert app.m.table.loc[1, "Dobro"] == 100


def test_formula_adulterada_no_estado_nao_executa_codigo(app, tmp_path):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")
    app.m.meta["formulas"]["Dobro"] = "__import__('os').system('echo pwned > pwned.txt')"

    app.m.commit()  # recalcula: tem de falhar sem executar nada

    assert not (tmp_path / "pwned.txt").exists()
    assert app.m.table["Dobro"].isna().all()


def test_formula_que_falha_avisa_e_deixa_coluna_em_branco(app, capsys):
    app.dummy()
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")
    app.m.meta["formulas"]["Dobro"] = "Apagada * 2"

    app.m.recalculate_formulas()

    assert "falhou" in capsys.readouterr().out and app.m.table["Dobro"].isna().all()
