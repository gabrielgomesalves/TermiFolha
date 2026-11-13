def _filtrar_texto(app, capsys, op, term="lisboa"):
    app.dummy()
    capsys.readouterr()
    app.run(app.m.filter_rows, "cidade", op, term)
    return capsys.readouterr().out


def test_filtro_de_texto_com_operador_invalido_avisa_e_nao_filtra(app, capsys):
    """Regressão M6: um operador desconhecido filtrava por 'contém' sem avisar."""
    saida = _filtrar_texto(app, capsys, "9")

    assert "Operador inválido" in saida
    assert "resultado(s)" not in saida


def test_filtro_de_texto_contem_ignora_maiusculas(app, capsys):
    saida = _filtrar_texto(app, capsys, "1", "LISBOA")

    assert "3 resultado(s)" in saida


def test_filtro_de_texto_igual_a(app, capsys):
    saida = _filtrar_texto(app, capsys, "2", "porto")

    assert "2 resultado(s)" in saida


def test_filtro_contem_trata_caracteres_especiais_como_texto_literal(app, capsys):
    """O termo não pode ser interpretado como expressão regular (`(` rebentava)."""
    app.m.table = app.m.pd.DataFrame({"Nota": ["a (b)", "c++", "plain"]})
    app.m.col_types = {"Nota": str}

    app.run(app.m.filter_rows, "nota", "1", "(")
    assert "1 resultado(s)" in capsys.readouterr().out

    app.run(app.m.filter_rows, "nota", "1", "c++")
    assert "1 resultado(s)" in capsys.readouterr().out
