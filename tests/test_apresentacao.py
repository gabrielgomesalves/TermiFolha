"""Apresentação: formatos de valores, colunas visíveis, páginas e linha de totais."""
import pandas as pd
import pytest


@pytest.mark.parametrize("fmt, esperado", [
    ({"decimals": 2, "thousands": True, "locale": "intl"}, "1,234.50"),
    ({"decimals": 2, "thousands": True, "locale": "pt"}, "1 234,50"),
    ({"decimals": 0, "thousands": False}, "1234"),  # 1234.5 arredonda para o par: 1234
    ({"decimals": 1, "prefix": "€", "suffix": "!"}, "€1234.5!"),
    ({"decimals": 99}, "1234.50"),  # inválido -> 2 casas
    ({"decimals": "x"}, "1234.50"),
])
def test_formato_numerico(app, fmt, esperado):
    assert app.m.apply_number_format(1234.5, fmt) == esperado


def test_format_cell_so_formata_numeros_validos(app):
    app.dummy()
    app.m.meta["formats"] = {"Salario": {"decimals": 1, "suffix": " €"}, "Nome": {"decimals": 1}}

    assert app.m.format_cell("Salario", 1500.5) == "1500.5 €"
    assert app.m.format_cell("Salario", float("nan")) == "--"
    assert app.m.format_cell("Nome", "João") == "João"  # texto nunca é formatado
    assert app.m.format_cell("Idade", 3) == "3"  # sem formato -> igual ao comportamento antigo


def test_formato_invalido_no_estado_nao_rebenta_o_desenho(app, capsys):
    app.dummy()
    app.m.meta["formats"] = {"Salario": {"decimals": 2, "prefix": 5}}  # prefixo com tipo errado

    app.m.draw_table()

    assert "João" in capsys.readouterr().out


def test_definir_formato_pela_interface(app, capsys):
    app.dummy()

    app.run(app.m.format_column, "Salario", "1", "1", "s", "2", "€ ", " EUR")

    assert app.m.meta["formats"]["Salario"] == {"decimals": 1, "thousands": True, "locale": "pt", "prefix": "€ ", "suffix": " EUR"}
    capsys.readouterr()
    app.m.draw_table()
    assert "€ 1 500,5 EUR" in capsys.readouterr().out
    assert app.m.table.loc[0, "Salario"] == 1500.50  # os dados não mudam


def test_remover_formato(app):
    app.dummy()
    app.m.meta["formats"]["Salario"] = {"decimals": 1}

    app.run(app.m.format_column, "Salario", "2")

    assert "Salario" not in app.m.meta["formats"]


def test_formatar_coluna_inteira_usa_zero_casas_por_omissao(app, capsys):
    app.dummy()

    app.run(app.m.format_column, "Idade", "1", "", "n", "1", "", "")

    assert app.m.meta["formats"]["Idade"]["decimals"] == 0
    capsys.readouterr()
    app.m.draw_table()
    assert " 25 " in capsys.readouterr().out


def test_colunas_visiveis_escolhem_e_ordenam(app, capsys):
    app.dummy()

    app.run(app.m.choose_visible_columns, "salario, nome")
    capsys.readouterr()
    app.m.draw_table()
    saida = capsys.readouterr().out

    assert app.m.visible_columns() == ["Salario", "Nome"]
    cabecalho = next(linha for linha in saida.splitlines() if "#" in linha and "|" in linha)
    assert cabecalho.index("Salario") < cabecalho.index("Nome") and "Cidade" not in cabecalho
    assert "3 coluna(s) oculta(s)" in saida
    assert list(app.m.table.columns) == ["Nome", "Idade", "Salario", "Cidade", "Admissao"]  # dados intactos


def test_colunas_visiveis_invalidas_nao_alteram(app):
    app.dummy()

    app.run(app.m.choose_visible_columns, "Nome, Inexistente")
    assert app.m.meta["visible"] == []
    app.run(app.m.choose_visible_columns, "Nome, nome")
    assert app.m.meta["visible"] == []


def test_enter_repoe_todas_as_colunas(app):
    app.dummy()
    app.m.meta["visible"] = ["Nome"]

    app.run(app.m.choose_visible_columns, "")

    assert app.m.meta["visible"] == [] and len(app.m.visible_columns()) == 5


def test_renomear_e_remover_coluna_actualizam_metadados(app):
    app.dummy()
    app.m.meta["visible"] = ["Nome", "Salario"]
    app.m.meta["formats"]["Salario"] = {"decimals": 1}
    app.m.meta["rules"]["Salario"] = {"min": 0}

    app.run(app.m.rename_column, "Salario", "Ordenado")
    assert app.m.meta["visible"] == ["Nome", "Ordenado"]
    assert "Ordenado" in app.m.meta["formats"] and "Ordenado" in app.m.meta["rules"] and "Salario" not in app.m.meta["formats"]

    app.run(app.m.remove_column, "Ordenado", "s")
    assert app.m.meta["visible"] == ["Nome"] and app.m.meta["formats"] == {} and app.m.meta["rules"] == {}


@pytest.mark.parametrize("total, tamanho, pagina, esperado", [
    (45, 20, 0, (0, 20, 3)),
    (45, 20, 2, (40, 45, 3)),
    (45, 20, 9, (40, 45, 3)),  # além do fim -> última
    (45, 20, -4, (0, 20, 3)),  # antes do início -> primeira
    (0, 20, 0, (0, 0, 1)),
    (20, 20, 0, (0, 20, 1)),
])
def test_limites_de_pagina(app, total, tamanho, pagina, esperado):
    assert app.m.page_bounds(total, tamanho, pagina) == esperado


def _folha_grande(app, linhas=45):
    app.m.table = pd.DataFrame({"N": range(linhas), "T": [f"t{i}" for i in range(linhas)]})
    app.m.col_types = {"N": int, "T": str}
    app.m.apply_types()


def test_ecra_principal_limita_linhas_e_avisa(app, capsys):
    _folha_grande(app, 45)

    app.m.draw_table()
    saida = capsys.readouterr().out

    assert "t29" in saida and "t30" not in saida
    assert "mais 15 linha(s)" in saida


def test_ver_por_paginas_navega_e_sai(app, capsys):
    _folha_grande(app, 45)

    app.run(app.m.view_pages, "n", "p", "3", "sair")
    saida = capsys.readouterr().out

    assert "PÁGINA 1/3" in saida and "PÁGINA 2/3" in saida and "PÁGINA 3/3" in saida
    assert "t44" in saida  # a última linha só aparece na página 3


def test_ver_por_paginas_respeita_tamanho_configurado(app, capsys):
    _folha_grande(app, 45)
    app.m.settings["page_size"] = 10

    app.run(app.m.view_pages, "sair")

    assert "PÁGINA 1/5" in capsys.readouterr().out


def test_linha_de_totais_so_na_apresentacao(app, capsys):
    app.dummy()

    app.run(app.m.choose_totals, "1")  # soma
    capsys.readouterr()
    app.m.draw_table()
    saida = capsys.readouterr().out

    assert app.m.meta["totals"] == "sum" and "Σ" in saida
    assert "16152" in saida
    assert len(app.m.table) == 8  # a linha de totais não entra nos dados


def test_valores_da_linha_de_totais(app):
    app.dummy()

    app.m.meta["totals"] = "sum"
    assert app.m.totals_row(["Nome", "Idade", "Salario"]) == {"Nome": "", "Idade": "221", "Salario": "16152"}

    app.m.meta["totals"] = "count"
    assert app.m.totals_row(["Nome", "Idade"]) == {"Nome": "8", "Idade": "7"}

    app.m.meta["totals"] = None
    assert app.m.totals_row(["Nome"]) is None


def test_totais_usam_o_formato_da_coluna(app):
    app.dummy()
    app.m.meta["formats"]["Salario"] = {"decimals": 0, "thousands": True, "locale": "intl"}
    app.m.meta["totals"] = "sum"

    assert app.m.totals_row(["Salario"]) == {"Salario": "16,152"}


def test_desligar_totais(app):
    app.dummy()
    app.m.meta["totals"] = "sum"

    app.run(app.m.choose_totals, "0")

    assert app.m.meta["totals"] is None


def test_definicoes_de_apresentacao_sobrevivem_ao_desfazer(app):
    app.dummy()
    app.run(app.m.choose_totals, "2")
    app.run(app.m.remove_row, "0", "s")

    app.m.undo_last()

    assert app.m.meta["totals"] == "mean" and len(app.m.table) == 8
