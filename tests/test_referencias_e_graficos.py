"""Fórmulas com referências a células e intervalos; título e rótulos configuráveis dos gráficos."""
import pytest


def _avaliar(app, expressao):
    return app.m.evaluate_formula(expressao, app.m.table, app.m.col_types)


def test_celula_por_letra_e_numero_da_linha(app):
    app.dummy()

    serie, tipo = _avaliar(app, "Salario / B0")  # B = Idade, linha 0 = primeira linha do ecrã

    assert serie[1] == pytest.approx(app.m.table.loc[1, "Salario"] / 25)
    assert _avaliar(app, "$B$0 + 1")[0].eq(26).all()
    assert _avaliar(app, "C1")[0].eq(app.m.table.loc[1, "Salario"]).all()
    assert _avaliar(app, "A0")[1] is str


def test_celula_por_nome_de_coluna(app):
    app.dummy()

    assert _avaliar(app, "[Salario]2")[0].eq(app.m.table.loc[2, "Salario"]).all()
    assert _avaliar(app, "[salario]2")[0].iloc[0] == app.m.table.loc[2, "Salario"]


def test_intervalos_e_agregados(app):
    app.dummy()
    t = app.m.table

    assert _avaliar(app, "soma(C0:C3)")[0].iloc[0] == pytest.approx(t["Salario"].iloc[0:4].sum())
    assert _avaliar(app, "SOMA(Salario)")[0].iloc[0] == pytest.approx(t["Salario"].sum())
    assert _avaliar(app, "media(Idade)")[0].iloc[0] == pytest.approx(t["Idade"].mean())  # ignora os vazios
    assert _avaliar(app, "contar(Idade)")[0].iloc[0] == t["Idade"].notna().sum()
    assert _avaliar(app, "min(C0:C7)")[0].iloc[0] == t["Salario"].min()
    assert _avaliar(app, "max(Salario)")[0].iloc[0] == t["Salario"].max()
    assert _avaliar(app, "max(Idade, 30)")[0][0] == 30  # com dois argumentos continua a ser linha a linha
    assert _avaliar(app, "soma(B0:C1)")[0].iloc[0] == pytest.approx(t[["Idade", "Salario"]].iloc[0:2].sum().sum())


def test_agregado_usado_numa_conta_por_linha(app):
    app.dummy()

    serie, tipo = _avaliar(app, "Salario / soma(Salario) * 100")

    assert tipo is float and serie.sum() == pytest.approx(100)


def test_intervalo_fora_de_um_agregado_e_recusado(app):
    app.dummy()

    for expressao in ("C0:C3", "C0:C3 + 1", "round(C0:C3)", "min(C0:C3, 1)"):
        with pytest.raises(app.m.FormulaError):
            _avaliar(app, expressao)


@pytest.mark.parametrize("expressao, trecho", [
    ("Z1 + 1", "não existe"), ("C99", "linha 99"), ("[Inexistente]1", "inexistente"), ("soma(C0:C99)", "linha 99"),
])
def test_referencias_invalidas_dao_erro_claro(app, expressao, trecho):
    app.dummy()

    with pytest.raises(app.m.FormulaError, match=trecho):
        _avaliar(app, expressao)


def test_texto_entre_aspas_e_parentesis_nao_sao_referencias(app):
    app.dummy()

    assert _avaliar(app, '"B2"')[0].eq("B2").all()
    assert app.m.canonical_formula('"B2" + [Preço 2]', ["A"]) == '"B2" + [Preço 2]'


def test_coluna_com_nome_de_celula_ganha_a_referencia(app):
    app.dummy()
    app.m.table["Q1"] = 7
    app.m.col_types["Q1"] = int

    serie, _ = _avaliar(app, "Q1 + 1")

    assert serie.eq(8).all()


def test_datas_em_celulas_e_intervalos(app):
    app.dummy()

    assert _avaliar(app, "year(E0)")[0].eq(2020).all()
    assert _avaliar(app, "min(E0:E7)")[1] == app.m.DATE_TYPE
    with pytest.raises(app.m.FormulaError, match="datas"):
        _avaliar(app, "soma(E0:E7)")


def test_letras_de_colunas(app):
    assert [app.m.column_letter(i) for i in (0, 25, 26, 27, 701, 702)] == ["A", "Z", "AA", "AB", "ZZ", "AAA"]


def test_formula_e_gravada_com_nomes_e_resiste_a_mudar_as_colunas_de_sitio(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Rel", "Salario / $C$0")
    app.run(app.m.add_formula_column, "Parte", "Salario / soma(C0:C7)")

    assert app.m.meta["formulas"]["Rel"] == "Salario / [Salario]0"
    assert app.m.meta["formulas"]["Parte"] == "Salario / soma([Salario]0:[Salario]7)"

    app.m.table = app.m.table[["Salario", "Nome", "Idade", "Cidade", "Admissao", "Rel", "Parte"]]  # Salario passa a ser a coluna A
    app.m.commit()

    assert app.m.table.loc[0, "Rel"] == pytest.approx(1.0)
    assert app.m.table["Parte"].sum() == pytest.approx(1.0)


def test_coluna_referida_por_uma_celula_nao_se_remove(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Rel", "Salario / C0")

    assert app.m.formula_users("Salario") == ["Rel"]
    assert app.m.formula_users("Nome") == []


def test_intervalo_conta_todas_as_colunas_entre_os_extremos(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Total", "soma(B0:C7)")

    assert app.m.formula_users("Idade") == ["Total"] and app.m.formula_users("Salario") == ["Total"]


def test_recalcula_quando_a_celula_referida_muda(app):
    app.dummy()
    app.run(app.m.add_formula_column, "Rel", "Salario / C0")

    app.run(app.m.edit_cell, "0", "Salario", "3000")

    assert app.m.table.loc[0, "Rel"] == pytest.approx(1.0)
    assert app.m.table.loc[1, "Rel"] == pytest.approx(app.m.table.loc[1, "Salario"] / 3000)


def test_referencia_a_linha_apagada_deixa_a_coluna_em_branco_e_avisa(app, capsys):
    app.dummy()
    app.run(app.m.add_formula_column, "Rel", "Salario / C7")
    app.run(app.m.remove_row, "7", "s")

    assert "falhou" in capsys.readouterr().out and app.m.table["Rel"].isna().all()


def test_ajuda_mostra_as_letras_das_colunas(app, capsys):
    app.dummy()

    app.run(app.m.add_formula_column, "sair")

    assert "A=Nome  B=Idade  C=Salario" in capsys.readouterr().out


# ---- título e rótulos dos gráficos ----

def test_titulo_e_rotulos_personalizados_aplicam_se_ao_proximo_grafico(app, tmp_path, monkeypatch):
    app.dummy()
    vistos = []
    monkeypatch.setattr(app.m.plt, "show", lambda: vistos.append([(a.get_title(), a.get_xlabel(), a.get_ylabel()) for a in app.m.plt.gcf().axes]))

    app.run(app.m.set_plot_labels, "Vendas 2026", "Pessoa", "Euros")
    app.run(app.m.plot_scatter, "Idade,Salario")
    app.run(app.m.plot_scatter, "Idade,Salario")

    assert vistos[0] == [("Vendas 2026", "Pessoa", "Euros")]
    assert vistos[1] == [("Dispersão: Idade vs Salario", "Idade", "Salario")]  # só valeu uma vez


def test_so_os_campos_preenchidos_mudam(app, monkeypatch):
    app.dummy()
    vistos = []
    monkeypatch.setattr(app.m.plt, "show", lambda: vistos.append(app.m.plt.gcf().axes[0].get_title()))

    app.run(app.m.set_plot_labels, "", "", "Y")
    app.run(app.m.plot_area, "Salario")

    assert vistos == ["Evolução da soma acumulada"]


def test_rotulos_personalizados_tambem_valem_ao_gravar_png(app, tmp_path, monkeypatch):
    app.dummy()
    monkeypatch.setattr(app.m, "plot_png_dir", str(tmp_path))
    titulos = []
    original = app.m.plt.savefig
    monkeypatch.setattr(app.m.plt, "savefig", lambda *a, **k: (titulos.append(app.m.plt.gcf().axes[0].get_title()), original(*a, **k)))

    app.run(app.m.set_plot_labels, "Titulo X", "", "")
    app.run(app.m.plot_area, "Salario")

    assert titulos == ["Titulo X"] and len(list(tmp_path.glob("area_*.png"))) == 1


def test_grafico_com_barra_de_cores_muda_o_eixo_principal(app, monkeypatch):
    app.dummy()
    vistos = []
    monkeypatch.setattr(app.m.plt, "show", lambda: vistos.append(app.m.plt.gcf().axes[0].get_title()))

    app.run(app.m.set_plot_labels, "Correlações", "", "")
    app.run(app.m.plot_correlation_heatmap, "Idade,Salario")

    assert vistos == ["Correlações"]


def test_texto_do_grafico_e_limpo(app):
    limpo = app.m.clean_plot_text("  Preço  em $ e $x$  ")
    assert limpo == r"Preço em \$ e \$x\$"
    assert len(app.m.clean_plot_text("x" * 500)) == app.m.PLOT_TEXT_MAX


def test_texto_com_cifrao_nao_rebenta_o_grafico(app, monkeypatch):
    app.dummy()
    monkeypatch.setattr(app.m.plt, "show", lambda: app.m.plt.gcf().canvas.draw())

    app.run(app.m.set_plot_labels, "Custo $ total $x^$", "", "")
    app.run(app.m.plot_area, "Salario")


def test_cancelar_dialogo_sem_texto_limpa_o_pendente(app):
    app.run(app.m.set_plot_labels, "A", "B", "C")
    app.run(app.m.set_plot_labels, "", "", "")

    assert not any(app.m.plot_custom.values())
