import pandas as pd
import pytest


@pytest.fixture
def xlsx_inteiros_com_falhas(tmp_path):
    path = tmp_path / "idades.xlsx"
    pd.DataFrame({"Idade": [25, None, 40], "Preco": [1.5, None, 2.0]}).to_excel(path, index=False)
    return str(path)


def test_importar_pede_confirmacao_se_ja_existe_folha(app, capsys):
    app.dummy()
    antes = app.m.table.copy()
    capsys.readouterr()

    app.run(app.m.import_excel, "n")  # recusa: nem chega a pedir o nome do ficheiro

    assert "Operação cancelada" in capsys.readouterr().out
    assert app.m.table.equals(antes)


def test_importar_csv_pede_confirmacao_se_ja_existe_folha(app, capsys):
    app.dummy()
    antes = app.m.table.copy()
    capsys.readouterr()

    app.run(app.m.import_csv, "n")

    assert "Operação cancelada" in capsys.readouterr().out
    assert app.m.table.equals(antes)


def test_importar_sem_folha_nao_pede_confirmacao(app, tmp_path):
    pd.DataFrame({"A": [1]}).to_excel(tmp_path / "a.xlsx", index=False)

    app.run(app.m.import_excel, "a")  # sem folha: vai direito ao nome

    assert list(app.m.table.columns) == ["A"]


def test_excel_inteiros_com_falhas_podem_ser_tratados_como_inteiro(app, xlsx_inteiros_com_falhas):
    """A coluna Idade só tem inteiros (e uma célula vazia); Preco tem decimais."""
    app.run(app.m.import_excel, xlsx_inteiros_com_falhas, "s")

    assert app.m.col_types["Idade"] is int
    assert str(app.m.table["Idade"].dtype) == "Int64"
    assert app.m.col_types["Preco"] is float  # tem 1.5: nunca é perguntada


def test_excel_inteiros_com_falhas_recusados_ficam_decimais(app, xlsx_inteiros_com_falhas):
    app.run(app.m.import_excel, xlsx_inteiros_com_falhas, "n")

    assert app.m.col_types["Idade"] is float


def test_excel_so_pergunta_quando_ha_ambiguidade(app, tmp_path):
    pd.DataFrame({"A": [1, 2], "B": [1.0, 2.5]}).to_excel(tmp_path / "x.xlsx", index=False)

    app.run(app.m.import_excel, "x")  # nenhuma pergunta extra: `feed` falharia se houvesse

    assert app.m.col_types == {"A": int, "B": float}


def test_csv_com_virgula_decimal_e_ponto_e_virgula(app, tmp_path):
    (tmp_path / "pt.csv").write_text("Preco;Qtd\n1,5;3\n2;4\n10,25;5\n", encoding="utf-8")

    app.run(app.m.import_csv, "pt")

    assert app.m.table["Preco"].tolist() == [1.5, 2.0, 10.25]
    assert app.m.col_types["Preco"] is float
    assert app.m.col_types["Qtd"] is int


def test_csv_com_virgula_decimal_entre_aspas(app, tmp_path):
    (tmp_path / "q.csv").write_text('Preco,Nome\n"1,5",a\n"2,25",b\n', encoding="utf-8")

    app.run(app.m.import_csv, "q")

    assert app.m.table["Preco"].tolist() == [1.5, 2.25]


def test_csv_texto_com_virgula_nao_e_confundido_com_numero(app, tmp_path):
    (tmp_path / "t.csv").write_text('Nota\n"1,5 kg"\n"2,0"\n', encoding="utf-8")

    app.run(app.m.import_csv, "t")

    assert app.m.col_types["Nota"] is str
    assert app.m.table["Nota"].tolist() == ["1,5 kg", "2,0"]


def test_exportar_csv_formato_excel_portugues_e_reimportavel(app, tmp_path):
    app.dummy()
    original = app.m.table.copy()

    app.run(app.m.export_csv, "pt", "2")  # [2] Excel PT: ';' e ','
    text = (tmp_path / "pt.csv").read_text(encoding="utf-8-sig")
    assert text.splitlines()[0] == "Nome;Idade;Salario;Cidade;Admissao"
    assert "1500,5" in text

    app.run(app.m.import_csv, "s", "pt")
    assert app.m.table["Salario"].tolist() == original["Salario"].tolist()
    assert app.m.table["Admissao"].tolist() == original["Admissao"].tolist()
    assert str(app.m.table["Idade"].dtype) == "Int64"


def test_exportar_csv_formato_internacional_por_omissao(app, tmp_path):
    app.dummy()

    app.run(app.m.export_csv, "int", "")

    assert "1500.5" in (tmp_path / "int.csv").read_text(encoding="utf-8-sig")


def test_estat_resumo_sem_coluna_descreve_todas_as_numericas(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.stat_summary, "")

    saida = capsys.readouterr().out
    assert "Idade" in saida.split("Resumo estatístico")[-1]
    assert "Salario" in saida.split("Resumo estatístico")[-1]
    assert "Nome" not in saida.split("Resumo estatístico")[-1]
    assert "Média" in saida


def test_estat_resumo_com_coluna_continua_a_funcionar(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.stat_summary, "salario")

    assert "Resumo estatístico de 'Salario'" in capsys.readouterr().out


def test_estat_unicos_sem_coluna_lista_todas(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.stat_unique, "")

    saida = capsys.readouterr().out
    import re
    assert re.search(r"Nome\s+7 \(de 8 não-nulos\)", saida)  # "João" repete-se
    assert re.search(r"Idade\s+6 \(de 7 não-nulos\)", saida)
    assert re.search(r"Cidade\s+5 \(de 8 não-nulos\)", saida)


def test_estat_unicos_com_coluna_continua_a_funcionar(app, capsys):
    app.dummy()
    capsys.readouterr()

    app.run(app.m.stat_unique, "cidade")

    assert "Valores únicos em 'Cidade': 5 (de 8 não-nulos)" in capsys.readouterr().out
