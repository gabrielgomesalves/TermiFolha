"""Fórmulas: referências relativas (anterior/seguinte/desloca) e referências a outras folhas."""
import json
from pathlib import Path

import pandas as pd
import pytest


def _livro(app):
    """Folha1 (dummy, activa), 'Vendas 2025' (Produto, Preço, Qtd; 3 linhas) e 'Taxas' (Taxa: 0.1 e 0.2)."""
    app.dummy()
    vendas = pd.DataFrame({"Produto": ["A", "B", "C"], "Preço": [10.0, 20.0, 30.0], "Qtd": pd.array([1, 2, 3], dtype="Int64")})
    app.m.add_sheet("Vendas 2025", app.m.typed_frame(vendas, {"Produto": str, "Preço": float, "Qtd": int}), {"Produto": str, "Preço": float, "Qtd": int}, activate=False)
    taxas = pd.DataFrame({"Taxa": [0.1, 0.2]})
    app.m.add_sheet("Taxas", taxas, {"Taxa": float}, activate=False)
    return app.m.table


def _avaliar(app, expressao):
    return app.m.evaluate_formula(expressao, app.m.table, app.m.col_types)


# ---------------------------------------------------------------- relativas

def test_anterior_e_seguinte_acompanham_a_linha(app):
    t = _livro(app)

    anterior, _ = _avaliar(app, "anterior(Salario)")
    seguinte, _ = _avaliar(app, "seguinte(Salario)")

    assert pd.isna(anterior[0]) and anterior[1:].tolist() == t["Salario"][:-1].tolist()
    assert pd.isna(seguinte.iloc[-1]) and seguinte[:-1].tolist() == t["Salario"][1:].tolist()


def test_desloca_com_n_positivo_e_negativo(app):
    t = _livro(app)

    assert _avaliar(app, "desloca(Salario, 2)")[0][:-2].tolist() == t["Salario"][2:].tolist()
    assert _avaliar(app, "desloca(Salario, -1)")[0].equals(_avaliar(app, "anterior(Salario)")[0])
    assert _avaliar(app, "desloca(Salario, 0)")[0].equals(t["Salario"].astype("float64"))


def test_anterior_com_varias_linhas(app):
    t = _livro(app)

    serie, _ = _avaliar(app, "anterior(Salario, 3)")

    assert serie[:3].isna().all() and serie[3:].tolist() == t["Salario"][:-3].tolist()


def test_variacao_face_a_linha_anterior(app):
    t = _livro(app)

    serie, tipo = _avaliar(app, "Salario - anterior(Salario)")

    assert tipo is float and pd.isna(serie[0])
    assert serie[1] == pytest.approx(t.loc[1, "Salario"] - t.loc[0, "Salario"])


def test_relativas_dentro_de_se(app):
    t = _livro(app)

    serie, _ = _avaliar(app, "se(Salario > anterior(Salario), 1, 0)")

    esperado = (t["Salario"][1:].to_numpy() > t["Salario"][:-1].to_numpy()).astype(int)
    assert serie[1:].tolist() == esperado.tolist()


def test_anterior_de_inteiros_continua_inteiro_com_celula_vazia(app):
    _livro(app)

    serie, tipo = _avaliar(app, "anterior(Idade)")

    assert tipo is int and pd.isna(serie[0])


def test_anterior_de_texto_e_de_datas(app):
    t = _livro(app)

    nomes, tipo_nome = _avaliar(app, "anterior(Nome)")
    datas, tipo_data = _avaliar(app, "anterior(Admissao)")

    assert tipo_nome is str and nomes[1] == t.loc[0, "Nome"] and pd.isna(nomes[0])
    assert tipo_data is app.m.DATE_TYPE and datas[1] == t.loc[0, "Admissao"]


def test_numero_de_linhas_tem_de_ser_inteiro_fixo(app):
    _livro(app)

    for expressao in ("anterior(Salario, 1.5)", "desloca(Salario, Idade)", "anterior(Salario, 'a')", "desloca(Salario, 99999999)"):
        with pytest.raises(app.m.FormulaError, match="inteiro fixo"):
            _avaliar(app, expressao)


def test_numero_de_argumentos_das_relativas(app):
    _livro(app)

    for expressao in ("anterior()", "seguinte(Salario, 1, 2)", "desloca(Salario)"):
        with pytest.raises(app.m.FormulaError, match="espera entre"):
            _avaliar(app, expressao)


def test_coluna_calculada_com_anterior_recalcula_ao_editar(app):
    t = _livro(app)
    app.run(app.m.add_formula_column, "Var", "Salario - anterior(Salario)")

    app.run(app.m.edit_cell, "0", "Salario", "1000")

    assert app.m.table.loc[1, "Var"] == pytest.approx(t.loc[1, "Salario"] - 1000)
    assert app.m.meta["formulas"]["Var"] == "Salario - anterior(Salario)"


# ---------------------------------------------------------------- outras folhas

def test_celula_de_outra_folha_por_nome_e_por_letra(app):
    _livro(app)

    assert _avaliar(app, "Taxas!A1")[0].eq(0.2).all()
    assert _avaliar(app, "Taxas![Taxa]0 + 1")[0].eq(1.1).all()
    assert _avaliar(app, "'Vendas 2025'!B2")[0].eq(30.0).all()  # B = Preço, linha 2
    assert _avaliar(app, "'Vendas 2025'![Preço]1 * 2")[0].eq(40.0).all()
    assert _avaliar(app, '"Vendas 2025"!C0')[0].eq(1).all()


def test_nome_da_folha_nao_distingue_maiusculas(app):
    _livro(app)

    assert _avaliar(app, "taxas!A0")[0].eq(0.1).all()
    assert _avaliar(app, "'vendas 2025'!B0")[0].eq(10.0).all()


def test_intervalo_e_coluna_inteira_de_outra_folha(app):
    t = _livro(app)

    assert _avaliar(app, "soma('Vendas 2025'!B0:B2)")[0].eq(60.0).all()
    assert _avaliar(app, "media('Vendas 2025'!B0:B1)")[0].eq(15.0).all()
    assert _avaliar(app, "soma('Vendas 2025'![Preço])")[0].eq(60.0).all()
    assert _avaliar(app, "contar('Vendas 2025'![Produto])")[0].eq(3).all()
    assert _avaliar(app, "max('Vendas 2025'![Qtd])")[0].eq(3).all()
    assert _avaliar(app, "Salario / soma(Taxas![Taxa])")[0][0] == pytest.approx(t.loc[0, "Salario"] / 0.30000000000000004)


def test_intervalo_de_varias_colunas_de_outra_folha(app):
    _livro(app)

    assert _avaliar(app, "soma('Vendas 2025'!B0:C1)")[0].eq(10 + 20 + 1 + 2).all()


def test_misturar_folhas_e_a_propria(app):
    t = _livro(app)

    serie, _ = _avaliar(app, "Salario * (1 + Taxas!A1) + Folha1!C0")

    assert serie[2] == pytest.approx(t.loc[2, "Salario"] * 1.2 + t.loc[0, "Salario"])


def test_erros_ao_referir_outras_folhas(app):
    _livro(app)

    with pytest.raises(app.m.FormulaError, match="folha inexistente: 'Nada'"):
        _avaliar(app, "Nada!A0")
    with pytest.raises(app.m.FormulaError, match="coluna inexistente em 'Taxas'"):
        _avaliar(app, "soma(Taxas![Outra])")
    with pytest.raises(app.m.FormulaError, match="a coluna C não existe"):
        _avaliar(app, "Taxas!C0")
    with pytest.raises(app.m.FormulaError, match="linha 5 não existe da folha 'Taxas'"):
        _avaliar(app, "Taxas!A5")
    with pytest.raises(app.m.FormulaError, match="só pode ser usada dentro de soma"):
        _avaliar(app, "Taxas![Taxa] * 2")
    with pytest.raises(app.m.FormulaError, match="intervalo só pode ser usado dentro de soma"):
        _avaliar(app, "Taxas!A0:A1 * 2")


def test_operador_diferente_nao_e_confundido_com_folha(app):
    t = _livro(app)

    serie, _ = _avaliar(app, "Idade != 25")

    assert serie[0] == 0 and serie[1] == int(t.loc[1, "Idade"] != 25)


def test_tipo_inteiro_vem_das_celulas_de_outra_folha(app):
    _livro(app)

    assert _avaliar(app, "'Vendas 2025'!C1 * 2")[1] is int
    assert _avaliar(app, "'Vendas 2025'!B1 * 2")[1] is float
    assert _avaliar(app, "soma('Vendas 2025'![Qtd])")[1] is int  # a soma de inteiros é inteira, como dentro da folha
    assert _avaliar(app, "media('Vendas 2025'![Qtd])")[1] is float


def test_forma_canonica_usa_nomes_e_aspas(app):
    _livro(app)
    colunas = list(app.m.table.columns)

    assert app.m.canonical_formula("'Vendas 2025'!B2 + taxas!A0", colunas) == "'Vendas 2025'![Preço]2 + 'Taxas'![Taxa]0"
    assert app.m.canonical_formula("soma('Vendas 2025'!B0:B2)", colunas) == "soma('Vendas 2025'![Preço]0:[Preço]2)"
    assert app.m.canonical_formula("soma('vendas 2025'![preço])", colunas) == "soma('Vendas 2025'![Preço])"
    canonica = app.m.canonical_formula("'Vendas 2025'!B2", colunas)
    assert app.m.canonical_formula(canonica, colunas) == canonica  # idempotente


def test_texto_entre_aspas_seguido_de_ponto_de_exclamacao_nao_e_folha_inexistente(app):
    _livro(app)

    serie, tipo = _avaliar(app, "'Olá!' + Nome")

    assert tipo is str and serie[0].startswith("Olá!")


def test_coluna_calculada_com_outra_folha_recalcula_quando_a_outra_muda(app):
    _livro(app)
    app.run(app.m.add_formula_column, "Total", "soma('Vendas 2025'![Preço])")
    assert app.m.table["Total"].eq(60.0).all()

    app.m.get_sheet("Vendas 2025")["table"].loc[0, "Preço"] = 100.0
    app.m.commit()

    assert app.m.table["Total"].eq(150.0).all()
    assert app.m.meta["formulas"]["Total"] == "soma('Vendas 2025'![Preço])"  # fórmula guardada em forma canónica


def test_formula_de_folha_inactiva_recalcula_quando_a_activa_muda(app):
    t = _livro(app)
    app.m.switch_sheet("Vendas 2025")
    app.run(app.m.add_formula_column, "Ref", "Folha1!C0 + Preço")
    app.m.switch_sheet("Folha1")

    app.run(app.m.edit_cell, "0", "Salario", "1000")

    assert app.m.get_sheet("Vendas 2025")["table"]["Ref"].tolist() == [1010.0, 1020.0, 1030.0]
    assert t is not None


def test_ordem_de_calculo_segue_as_dependencias_entre_folhas(app):
    _livro(app)
    # 'Vendas 2025' tem uma coluna calculada; a Folha1 (que vem antes no livro) depende dela
    app.m.switch_sheet("Vendas 2025")
    app.run(app.m.add_formula_column, "Total", "Preço * Qtd")
    app.m.switch_sheet("Folha1")
    app.run(app.m.add_formula_column, "Soma", "soma('Vendas 2025'![Total])")
    assert app.m.table["Soma"].eq(10 + 40 + 90).all()

    app.m.switch_sheet("Vendas 2025")
    app.run(app.m.edit_cell, "0", "Qtd", "5")

    app.m.switch_sheet("Folha1")
    assert app.m.table["Soma"].eq(50 + 40 + 90).all()


def test_ciclo_entre_folhas_deixa_as_colunas_em_branco_e_avisa(app, capsys):
    _livro(app)
    app.m.table["X"] = 1.0
    app.m.col_types["X"] = float
    app.m.meta["formulas"]["X"] = "soma('Vendas 2025'![Y])"
    vendas = app.m.get_sheet("Vendas 2025")
    vendas["table"]["Y"] = 1.0
    vendas["types"]["Y"] = float
    vendas["meta"]["formulas"]["Y"] = "soma('Folha1'![X])"

    app.m.recalculate_formulas()

    saida = capsys.readouterr().out
    assert "referência circular" in saida and "Vendas 2025!Y" in saida and "'X'" in saida
    assert app.m.table["X"].isna().all() and vendas["table"]["Y"].isna().all()


def test_formula_que_se_refere_a_si_propria_e_circular(app, capsys):
    _livro(app)
    app.m.table["X"] = 1.0
    app.m.col_types["X"] = float
    app.m.meta["formulas"]["X"] = "soma(Folha1![X])"

    app.m.recalculate_formulas()

    assert "referência circular" in capsys.readouterr().out and app.m.table["X"].isna().all()


def test_folha_que_deixou_de_existir_avisa_sem_rebentar(app, capsys):
    _livro(app)
    app.m.table["X"] = 1.0
    app.m.col_types["X"] = float
    app.m.meta["formulas"]["X"] = "Fantasma!A0"

    app.m.recalculate_formulas()

    assert "folha inexistente" in capsys.readouterr().out and app.m.table["X"].isna().all()


def test_coluna_usada_por_outra_folha_nao_se_remove_nem_renomeia(app, capsys):
    _livro(app)
    app.m.switch_sheet("Vendas 2025")
    app.run(app.m.add_formula_column, "Ref", "Folha1!C0 + Preço")
    app.m.switch_sheet("Folha1")

    assert app.m.formula_users("Salario") == ["Vendas 2025!Ref"]
    app.run(app.m.remove_column, "Salario")
    app.run(app.m.rename_column, "Salario")

    saida = capsys.readouterr().out
    assert saida.count("usada pela fórmula de: Vendas 2025!Ref") == 2
    assert "Salario" in app.m.table.columns


def test_folha_usada_por_formulas_nao_se_elimina(app, capsys):
    _livro(app)
    app.run(app.m.add_formula_column, "Total", "soma(Taxas![Taxa])")

    with pytest.raises(ValueError, match="usada pelas fórmulas de: Folha1!Total"):
        app.m.delete_sheet("Taxas")
    app.run(app.m.delete_sheet_ui, "Taxas", "s")

    assert "Taxas" in app.m.sheet_names() and "usada pelas fórmulas" in capsys.readouterr().out
    app.run(app.m.freeze_formula, "Total", "s")  # depois de converter em valores já se pode eliminar
    app.m.delete_sheet("Taxas")
    assert "Taxas" not in app.m.sheet_names()


def test_a_propria_folha_pode_referir_se_pelo_nome_e_eliminar_outras(app):
    _livro(app)
    app.run(app.m.add_formula_column, "Eco", "Folha1!C0")

    app.m.delete_sheet("Taxas")  # ninguém usa a Taxas

    assert app.m.table["Eco"].eq(app.m.table.loc[0, "Salario"]).all()
    assert app.m.sheet_formula_users("Folha1") == []  # só conta fórmulas de OUTRAS folhas


def test_renomear_folha_actualiza_as_formulas_que_a_referem(app):
    _livro(app)
    app.run(app.m.add_formula_column, "Total", "soma('Vendas 2025'![Preço]) + Taxas!A0")

    app.m.rename_sheet("Vendas 2025", "Vendas")
    app.m.rename_sheet("Taxas", "O'Brien")
    app.m.commit()

    assert app.m.meta["formulas"]["Total"] == "soma('Vendas'![Preço]) + \"O'Brien\"![Taxa]0"
    assert app.m.table["Total"].eq(60.1).all()


def test_renomear_folha_inactiva_actualiza_formulas_de_outra_inactiva(app):
    _livro(app)
    app.m.switch_sheet("Vendas 2025")
    app.run(app.m.add_formula_column, "Ref", "Taxas!A0 + Preço")
    app.m.switch_sheet("Folha1")

    app.m.rename_sheet("Taxas", "Impostos")

    assert app.m.get_sheet("Vendas 2025")["meta"]["formulas"]["Ref"] == "'Impostos'![Taxa]0 + Preço"


def test_nome_de_folha_com_os_dois_tipos_de_aspas_e_recusado(app):
    app.dummy()

    assert "aspas simples e duplas" in app.m.sheet_name_error("a'b\"c")
    assert app.m.sheet_name_error("a'b") is None and app.m.sheet_name_error('a"b') is None


def test_formulas_entre_folhas_gravam_se_e_recarregam_se(app):
    _livro(app)
    app.run(app.m.add_formula_column, "Total", "soma('Vendas 2025'![Preço]) + Taxas!A0")
    estado = json.loads(Path(app.m.STATE_FILE).read_text(encoding="utf-8"))
    assert estado["folhas"][0]["formulas"]["Total"] == "soma('Vendas 2025'![Preço]) + 'Taxas'![Taxa]0"

    app.m.reset_workbook()
    app.m.load_state()

    assert app.m.table["Total"].eq(60.1).all()


def test_desfazer_repoe_a_folha_e_a_formula(app):
    _livro(app)
    app.run(app.m.add_formula_column, "Total", "soma(Taxas![Taxa])")

    app.run(app.m.undo_last)

    assert "Total" not in app.m.table.columns and app.m.meta["formulas"] == {}


def test_a_ajuda_das_formulas_fala_de_relativas_e_outras_folhas(app, capsys):
    _livro(app)

    app.run(app.m.add_formula_column, "x", "sair")

    saida = capsys.readouterr().out
    assert "anterior(x)" in saida and "Outras folhas:" in saida and "Folhas:   Vendas 2025  Taxas" in saida
