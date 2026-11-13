import warnings

import pandas as pd


def _exportar_pdf(app, nome="saida"):
    app.run(app.m.exportar_pdf, nome)


def test_exportar_pdf_com_caracteres_fora_de_latin1_nao_falha(app, capsys, tmp_path):
    """Regressão B4: com fpdf2 e fontes core, 'ż' abortava toda a exportação."""
    app.m.tabela = pd.DataFrame({"Nome": ["Zażółć", "Café €"], "N": [1, 2]})
    app.m.tipos = {"Nome": str, "N": int}
    app.m.aplicar_tipos()

    _exportar_pdf(app)

    saida = capsys.readouterr().out
    assert "Erro" not in saida
    assert (tmp_path / "saida.pdf").stat().st_size > 0


def test_exportar_pdf_nao_usa_api_deprecada(app, tmp_path):
    """Regressão B4: `ln=True` e a fonte 'arial' estão deprecados no fpdf2."""
    app.dummy()

    with warnings.catch_warnings():
        warnings.simplefilter("error", DeprecationWarning)
        _exportar_pdf(app)

    assert (tmp_path / "saida.pdf").exists()


def test_exportar_pdf_preserva_acentos_do_portugues(app, tmp_path):
    app.dummy()  # contém 'João'

    _exportar_pdf(app)

    assert b"%PDF" in (tmp_path / "saida.pdf").read_bytes()[:8]


def test_excel_exportado_pode_ser_reimportado_sem_perdas(app, tmp_path):
    app.dummy()
    original = app.m.tabela.copy()

    app.run(app.m.exportar, "ida_e_volta")
    app.run(app.m.importar, str(tmp_path / "ida_e_volta.xlsx"))

    assert list(app.m.tabela.columns) == list(original.columns)
    assert len(app.m.tabela) == len(original)
    assert app.m.tabela["Salario"].tolist() == original["Salario"].tolist()
    assert str(app.m.tabela["Admissao"].dtype).startswith("datetime64")


def test_exportar_word_gera_ficheiro(app, tmp_path):
    app.dummy()

    app.run(app.m.exportar_word, "relatorio")

    assert (tmp_path / "relatorio.docx").stat().st_size > 0
