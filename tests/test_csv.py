import pandas as pd


def test_csv_ida_e_volta_preserva_dados_e_tipos(app, tmp_path):
    app.dummy()
    original = app.m.tabela.copy()

    app.run(app.m.exportar_csv, "folha", "")
    ficheiro = tmp_path / "folha.csv"
    assert ficheiro.exists()

    app.run(app.m.importar_csv, str(ficheiro))

    t = app.m.tabela
    assert list(t.columns) == list(original.columns)
    assert len(t) == len(original)
    assert str(t["Idade"].dtype) == "Int64"
    assert int(t["Idade"].isna().sum()) == 1
    assert t["Salario"].tolist() == original["Salario"].tolist()
    assert pd.api.types.is_datetime64_any_dtype(t["Admissao"])
    assert t["Admissao"].tolist() == original["Admissao"].tolist()
    assert t["Nome"].tolist() == original["Nome"].tolist()  # acentos (João) intactos
    assert app.m.tipos["Admissao"] == app.m.TIPO_DATA


def test_csv_exportado_usa_datas_em_formato_portugues(app, tmp_path):
    app.dummy()

    app.run(app.m.exportar_csv, "folha", "")

    assert "15/01/2020" in (tmp_path / "folha.csv").read_text(encoding="utf-8-sig")


def test_importar_csv_com_ponto_e_virgula(app, tmp_path):
    (tmp_path / "pt.csv").write_text("Produto;Qtd\nCafé;3\nChá;5\n", encoding="utf-8")

    app.run(app.m.importar_csv, "pt")  # extensão acrescentada automaticamente

    assert list(app.m.tabela.columns) == ["Produto", "Qtd"]
    assert app.m.tabela["Qtd"].tolist() == [3, 5]
    assert app.m.tipos["Qtd"] is int


def test_importar_csv_com_uma_so_coluna(app, tmp_path):
    (tmp_path / "um.csv").write_text("Nome\nAna\nRui\n", encoding="utf-8")

    app.run(app.m.importar_csv, "um")

    assert list(app.m.tabela.columns) == ["Nome"]
    assert len(app.m.tabela) == 2


def test_importar_csv_em_latin1(app, tmp_path):
    (tmp_path / "antigo.csv").write_bytes("Nome;Cidade\nJoão;Évora\n".encode("latin-1"))

    app.run(app.m.importar_csv, "antigo")

    assert app.m.tabela.loc[0, "Cidade"] == "Évora"


def test_importar_csv_inexistente_nao_altera_a_folha(app, capsys):
    app.dummy()
    antes = app.m.tabela.copy()

    app.run(app.m.importar_csv, "s", "nao_existe")

    assert "não encontrado" in capsys.readouterr().out
    assert app.m.tabela.equals(antes)


def test_importar_csv_com_cabecalhos_numericos(app, tmp_path):
    (tmp_path / "anos.csv").write_text("2020,2021\n1,2\n3,4\n", encoding="utf-8")

    app.run(app.m.importar_csv, "anos")

    assert list(app.m.tabela.columns) == ["2020", "2021"]


def test_importar_csv_vazio_avisa_e_nao_altera(app, tmp_path, capsys):
    (tmp_path / "vazio.csv").write_text("", encoding="utf-8")
    app.dummy()
    linhas = len(app.m.tabela)

    app.run(app.m.importar_csv, "s", "vazio")

    assert "Erro" in capsys.readouterr().out
    assert len(app.m.tabela) == linhas


def test_exportar_csv_sem_folha_avisa(app, capsys):
    app.m.exportar_csv()

    assert "Sem dados para exportar" in capsys.readouterr().out


def test_importar_excel_continua_a_funcionar_apos_refactor(app, tmp_path):
    pd.DataFrame({"A": [1, 2], "B": ["x", "y"]}).to_excel(tmp_path / "e.xlsx", index=False)

    app.run(app.m.importar, "e")

    assert app.m.tipos == {"A": int, "B": str}


def test_importar_csv_separado_por_tabulador(app, tmp_path):
    (tmp_path / "tab.csv").write_text("A\tB\n1\tx\n2\ty\n", encoding="utf-8")

    app.run(app.m.importar_csv, "tab")

    assert list(app.m.tabela.columns) == ["A", "B"]


def test_importar_csv_preserva_zeros_a_esquerda(app, tmp_path):
    (tmp_path / "cod.csv").write_text("Codigo,Qtd\n007,1\n042,2\n", encoding="utf-8")

    app.run(app.m.importar_csv, "cod")

    assert app.m.tabela["Codigo"].tolist() == ["007", "042"]
    assert app.m.tipos["Codigo"] is str
    assert app.m.tipos["Qtd"] is int


def test_menu_folha_expoe_importar_e_exportar_csv(app, tmp_path):
    (tmp_path / "dados.csv").write_text("A,B\n1,x\n", encoding="utf-8")

    app.run(app.m.menu_folha, "8", "dados", "9", "copia", "", "0")

    assert list(app.m.tabela.columns) == ["A", "B"]
    assert (tmp_path / "copia.csv").exists()
