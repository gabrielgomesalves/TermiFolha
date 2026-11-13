"""Menu de folhas, juntar (join), empilhar, tabela dinâmica e Excel com várias folhas."""
import pandas as pd
import pytest


def _clientes_e_pedidos(app):
    """Folha activa 'Clientes' e outra 'Pedidos' (3 clientes; pedidos de 1, 1, 2 e 9 — o 9 não existe nos clientes)."""
    app.m.table = pd.DataFrame({"Id": [1, 2, 3], "Nome": ["Ana", "Rui", "Eva"]})
    app.m.col_types = {"Id": int, "Nome": str}
    app.m.apply_types()
    app.m.sheets = {"Clientes": None}
    app.m.active_name = "Clientes"
    pedidos = pd.DataFrame({"Id": [1, 1, 2, 9], "Valor": [10.0, 20.0, 5.0, 7.0], "Qtd": [1, 2, 3, 4], "Nome": ["a", "b", "c", "d"]})
    app.m.add_sheet("Pedidos", pedidos, {"Id": int, "Valor": float, "Qtd": int, "Nome": str}, activate=False)
    app.m.undo_stack.clear()
    app.m.undo_stack.append(app.m.snapshot())


@pytest.mark.parametrize("how, linhas", [("inner", 3), ("left", 4), ("right", 4), ("outer", 5)])
def test_juntar_numero_de_linhas_por_tipo(app, how, linhas):
    _clientes_e_pedidos(app)

    juntada, _ = app.m.merge_tables("Clientes", "Pedidos", "Id", "Id", how)

    assert len(juntada) == linhas


def test_juntar_renomeia_colunas_repetidas_e_preserva_tipos(app):
    _clientes_e_pedidos(app)

    juntada, tipos = app.m.merge_tables("Clientes", "Pedidos", "Id", "Id", "left")

    assert list(juntada.columns) == ["Id", "Nome", "Valor", "Qtd", "Nome_Pedidos"]
    assert tipos == {"Id": int, "Nome": str, "Valor": float, "Qtd": int, "Nome_Pedidos": str}
    assert str(juntada["Qtd"].dtype) == "Int64" and int(juntada["Qtd"].isna().sum()) == 1  # Eva sem pedidos
    assert juntada["Valor"].tolist()[:3] == [10.0, 20.0, 5.0]


def test_juntar_com_chaves_de_nomes_diferentes_mantem_as_duas(app):
    _clientes_e_pedidos(app)
    app.m.sheets["Pedidos"]["table"] = app.m.sheets["Pedidos"]["table"].rename(columns={"Id": "ClienteId"})
    app.m.sheets["Pedidos"]["types"]["ClienteId"] = app.m.sheets["Pedidos"]["types"].pop("Id")

    juntada, _ = app.m.merge_tables("Clientes", "Pedidos", "Id", "ClienteId", "inner")

    assert "Id" in juntada.columns and "ClienteId" in juntada.columns and len(juntada) == 3


def test_juntar_nao_altera_as_folhas_originais(app):
    _clientes_e_pedidos(app)

    app.m.merge_tables("Clientes", "Pedidos", "Id", "Id", "outer")

    assert list(app.m.table.columns) == ["Id", "Nome"] and len(app.m.sheets["Pedidos"]["table"]) == 4


def test_juntar_chaves_de_tipos_incompativeis_levanta(app):
    _clientes_e_pedidos(app)

    with pytest.raises(ValueError, match="tipos diferentes"):
        app.m.merge_tables("Clientes", "Pedidos", "Nome", "Id", "inner")


def test_juntar_inteiro_com_decimal_e_permitido(app):
    _clientes_e_pedidos(app)

    juntada, tipos = app.m.merge_tables("Clientes", "Pedidos", "Id", "Valor", "outer")

    assert len(juntada) > 0 and tipos["Id"] is float and tipos["Valor"] is float


def test_juntar_limita_explosao_de_linhas(app, monkeypatch):
    _clientes_e_pedidos(app)
    monkeypatch.setattr(app.m, "MAX_MERGE_ROWS", 5)

    with pytest.raises(ValueError, match="demasiadas linhas"):
        app.m.merge_tables("Clientes", "Pedidos", "Id", "Id", "inner")


def test_juntar_pela_interface_cria_folha_nova_e_activa(app):
    _clientes_e_pedidos(app)

    app.run(app.m.merge_sheets_ui, "Pedidos", "Id", "Id", "2")

    assert app.m.active_name == "Clientes+Pedidos" and len(app.m.table) == 4
    assert app.m.sheet_names() == ["Clientes", "Pedidos", "Clientes+Pedidos"]


@pytest.mark.parametrize("respostas", [
    ("sair",),
    ("Pedidos", "Inexistente", "Id"),
    ("Pedidos", "Id", "Inexistente"),
    ("Pedidos", "Id", "Id", "9"),
    ("Pedidos", "Nome", "Id", "1"),  # tipos incompatíveis
])
def test_juntar_pela_interface_cancelado_ou_invalido_nao_cria_folha(app, respostas):
    _clientes_e_pedidos(app)

    app.run(app.m.merge_sheets_ui, *respostas)

    assert app.m.sheet_names() == ["Clientes", "Pedidos"]


def test_juntar_precisa_de_duas_folhas(app, capsys):
    app.dummy()

    app.m.merge_sheets_ui()

    assert "pelo menos duas" in capsys.readouterr().out


# ---------------------------------------------------------------- empilhar


def test_empilhar_une_colunas_sem_distinguir_maiusculas(app):
    _clientes_e_pedidos(app)

    empilhada, tipos, a_texto = app.m.concat_tables("Clientes", "Pedidos")

    assert list(empilhada.columns) == ["Id", "Nome", "Valor", "Qtd"] and len(empilhada) == 7
    assert a_texto == [] and tipos["Qtd"] is int
    assert int(empilhada["Qtd"].isna().sum()) == 3 and int(empilhada["Valor"].isna().sum()) == 3  # colunas em falta ficam em branco
    assert str(empilhada["Qtd"].dtype) == "Int64"


def test_empilhar_tipos_diferentes_inteiro_e_decimal_dao_decimal(app):
    _clientes_e_pedidos(app)
    app.m.sheets["Pedidos"]["types"]["Id"] = float
    app.m.sheets["Pedidos"]["table"]["Id"] = app.m.sheets["Pedidos"]["table"]["Id"].astype(float)

    _, tipos, a_texto = app.m.concat_tables("Clientes", "Pedidos")

    assert tipos["Id"] is float and a_texto == []


def test_empilhar_tipos_incompativeis_passam_a_texto(app):
    _clientes_e_pedidos(app)
    app.m.sheets["Pedidos"]["table"]["Id"] = app.m.sheets["Pedidos"]["table"]["Id"].astype(str)
    app.m.sheets["Pedidos"]["types"]["Id"] = str

    empilhada, tipos, a_texto = app.m.concat_tables("Clientes", "Pedidos")

    assert a_texto == ["Id"] and tipos["Id"] is str
    assert empilhada["Id"].tolist()[:3] == ["1", "2", "3"]  # inteiros sem '.0'


def test_empilhar_pela_interface(app):
    _clientes_e_pedidos(app)

    app.run(app.m.concat_sheets_ui, "2")  # folha nº 2 = Pedidos

    assert app.m.active_name == "Clientes+Pedidos" and len(app.m.table) == 7


def test_empilhar_com_aviso_de_texto_pode_ser_cancelado(app):
    _clientes_e_pedidos(app)
    app.m.sheets["Pedidos"]["types"]["Id"] = str
    app.m.sheets["Pedidos"]["table"]["Id"] = app.m.sheets["Pedidos"]["table"]["Id"].astype(str)

    app.run(app.m.concat_sheets_ui, "Pedidos", "n")

    assert app.m.sheet_names() == ["Clientes", "Pedidos"]


# ---------------------------------------------------------------- menu de folhas


def test_escolher_folha_por_numero_e_por_nome(app):
    _clientes_e_pedidos(app)

    app.run(app.m.switch_sheet_ui, "2")
    assert app.m.active_name == "Pedidos"

    app.run(app.m.switch_sheet_ui, "clientes")
    assert app.m.active_name == "Clientes"

    app.run(app.m.switch_sheet_ui, "Nenhuma")
    app.run(app.m.switch_sheet_ui, "sair")
    assert app.m.active_name == "Clientes"


def test_cabecalho_mostra_a_folha_activa_quando_ha_varias(app, capsys):
    _clientes_e_pedidos(app)

    app.m.header("X")

    assert "folha 'Clientes' (1/2)" in capsys.readouterr().out


def test_adicionar_folha_em_branco_pela_interface(app):
    _clientes_e_pedidos(app)

    app.run(app.m.add_blank_sheet, "Notas", "Texto", "1", "")

    assert app.m.active_name == "Notas" and list(app.m.table.columns) == ["Texto"]
    assert app.m.sheet_names() == ["Clientes", "Pedidos", "Notas"]


def test_adicionar_folha_cancelando_a_definicao_nao_deixa_folha_vazia(app):
    _clientes_e_pedidos(app)

    app.run(app.m.add_blank_sheet, "Notas", "sair")

    assert app.m.sheet_names() == ["Clientes", "Pedidos"] and app.m.active_name == "Clientes"


def test_adicionar_folha_com_nome_repetido_pede_outro(app):
    _clientes_e_pedidos(app)

    app.run(app.m.add_blank_sheet, "clientes", "Notas", "Texto", "1", "")

    assert app.m.sheet_names() == ["Clientes", "Pedidos", "Notas"]


def test_renomear_eliminar_e_duplicar_pela_interface(app):
    _clientes_e_pedidos(app)

    app.run(app.m.rename_sheet_ui, "Pedidos", "Compras")
    assert app.m.sheet_names() == ["Clientes", "Compras"]

    app.run(app.m.duplicate_sheet_ui, "Compras", "Compras2")
    assert app.m.active_name == "Compras2" and len(app.m.table) == 4

    app.run(app.m.delete_sheet_ui, "Compras2", "n")
    assert "Compras2" in app.m.sheet_names()
    app.run(app.m.delete_sheet_ui, "Compras2", "s")
    assert app.m.sheet_names() == ["Clientes", "Compras"] and app.m.active_name == "Compras"


def test_eliminar_a_unica_folha_e_recusado(app, capsys):
    app.dummy()

    app.m.delete_sheet_ui()

    assert "pelo menos uma" in capsys.readouterr().out


def test_eliminar_folha_pode_ser_desfeito(app):
    _clientes_e_pedidos(app)
    app.run(app.m.delete_sheet_ui, "Pedidos", "s")

    app.m.undo_last()

    assert app.m.sheet_names() == ["Clientes", "Pedidos"] and len(app.m.sheets["Pedidos"]["table"]) == 4


def test_menu_de_folhas_navega_e_sai(app, capsys):
    _clientes_e_pedidos(app)

    app.run(app.m.menu_workbook, "m", "2", "0")

    assert app.m.active_name == "Pedidos" and "(4 linhas × 4 colunas)" in capsys.readouterr().out


# ---------------------------------------------------------------- tabela dinâmica


def test_tabela_dinamica_simples(app):
    app.dummy()

    resultado, tipos = app.m.pivot_data("Cidade", None, "Salario", "sum")

    assert list(resultado.columns) == ["Cidade", "Salario"]
    assert dict(zip(resultado["Cidade"], resultado["Salario"], strict=True))["Lisboa"] == 1500.5 + 3500.0 + 1500.5
    assert tipos == {"Cidade": str, "Salario": float}


def test_tabela_dinamica_com_colunas_e_totais(app):
    app.dummy()

    resultado, _ = app.m.pivot_data("Cidade", "Nome", "Salario", "sum", totals=True)

    assert "Total" in resultado.columns and resultado["Cidade"].tolist()[-1] == "Total"
    assert resultado.loc[resultado["Cidade"] == "Total", "Total"].iloc[0] == pytest.approx(16152.0)


def test_tabela_dinamica_contagem_e_media(app):
    app.dummy()

    contagem, _ = app.m.pivot_data("Cidade", None, "Idade", "count")
    media, _ = app.m.pivot_data("Cidade", None, "Idade", "mean")

    assert dict(zip(contagem["Cidade"], contagem["Idade"], strict=True))["Lisboa"] == 3
    assert dict(zip(media["Cidade"], media["Idade"], strict=True))["Porto"] == 29


def test_tabela_dinamica_com_categorias_numericas_gera_nomes_unicos(app):
    app.m.table = pd.DataFrame({"G": ["a", "a", "b"], "N": [1, 2, 1], "V": [1.0, 2.0, 3.0]})
    app.m.col_types = {"G": str, "N": int, "V": float}
    app.m.apply_types()

    resultado, _ = app.m.pivot_data("G", "N", "V", "sum")

    assert list(resultado.columns) == ["G", "1", "2"]


def test_tabela_dinamica_papeis_repetidos_ou_funcao_invalida(app):
    app.dummy()

    with pytest.raises(ValueError, match="diferentes"):
        app.m.pivot_data("Cidade", "Cidade", "Salario", "sum")
    with pytest.raises(ValueError, match="função desconhecida"):
        app.m.pivot_data("Cidade", None, "Salario", "median")


def test_tabela_dinamica_pela_interface_guarda_nova_folha(app):
    app.dummy()

    app.run(app.m.pivot_table_ui, "Cidade", "", "Salario", "2", "n", "s")

    assert app.m.active_name == "Dinamica" and list(app.m.table.columns) == ["Cidade", "Salario"]
    assert app.m.sheet_names() == ["Folha1", "Dinamica"]


def test_tabela_dinamica_sem_guardar_nao_cria_folha(app):
    app.dummy()

    app.run(app.m.pivot_table_ui, "Cidade", "Nome", "Salario", "1", "s", "n")

    assert app.m.sheet_names() == ["Folha1"]


# ---------------------------------------------------------------- Excel com várias folhas


def test_exportar_e_importar_excel_com_varias_folhas(app, tmp_path):
    _clientes_e_pedidos(app)
    caminho = str(tmp_path / "livro.xlsx")

    app.run(app.m.export_excel, caminho, "s")
    folhas = pd.read_excel(caminho, sheet_name=None)
    assert list(folhas) == ["Clientes", "Pedidos"] and len(folhas["Pedidos"]) == 4

    app.m.reset_workbook()
    app.run(app.m.import_excel, caminho)  # sem folha activa: não pede confirmação

    assert app.m.sheet_names() == ["Clientes", "Pedidos"] and app.m.active_name == "Clientes"
    assert str(app.m.table["Id"].dtype) == "Int64" and app.m.col_types["Id"] is int
    assert len(app.m.sheets["Pedidos"]["table"]) == 4


def test_exportar_excel_so_a_folha_activa(app, tmp_path):
    _clientes_e_pedidos(app)
    caminho = str(tmp_path / "um.xlsx")

    app.run(app.m.export_excel, caminho, "n")

    assert list(pd.read_excel(caminho, sheet_name=None)) == ["Sheet1"]


def test_importar_excel_de_uma_folha_continua_a_substituir_so_a_activa(app, tmp_path):
    _clientes_e_pedidos(app)
    caminho = tmp_path / "simples.xlsx"
    pd.DataFrame({"A": [1, 2]}).to_excel(caminho, index=False)

    app.run(app.m.import_excel, "s", str(caminho))

    assert app.m.sheet_names() == ["Clientes", "Pedidos"] and list(app.m.table.columns) == ["A"]
