"""Comportamentos que o manual promete e o código não cumpria."""


def test_remover_linha_mostra_o_conteudo_antes_de_confirmar(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.remove_row, "4", "n")

    # A linha 4 da folha dummy é o Carlos. O cabeçalho do ecrã já lista a tabela
    # (1 ocorrência); a segunda é a linha mostrada antes da confirmação.
    assert capsys.readouterr().out.count("Carlos") == 2


def test_amostra_por_percentagem(app, capsys):
    app.dummy()  # 8 linhas
    capsys.readouterr()

    app.run(app.m.random_sample, "50%")

    assert "Amostra aleatória de 4 linha(s)" in capsys.readouterr().out


def test_amostra_por_numero_fixo(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.random_sample, "3")

    assert "Amostra aleatória de 3 linha(s)" in capsys.readouterr().out


def test_amostra_valor_invalido_usa_predefinicao(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.random_sample, "abc")

    assert "Amostra aleatória de 5 linha(s)" in capsys.readouterr().out


def test_amostra_percentagem_fora_do_intervalo_usa_predefinicao(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.random_sample, "150%")

    assert "Amostra aleatória de 5 linha(s)" in capsys.readouterr().out
