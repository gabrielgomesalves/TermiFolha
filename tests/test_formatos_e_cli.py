"""JSON, texto colado, Markdown, HTML, projectos, definições e modo de linha de comando."""
import json
import os
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parent.parent

# ---------------------------------------------------------------- JSON


def test_json_lista_de_objectos_infere_tipos_e_preserva_inteiros_com_nulos(app):
    dados = [
        {"Nome": "Ana", "Idade": 30, "Nota": 1.5, "Ok": True},
        {"Nome": "Rui", "Idade": None, "Nota": 2, "Ok": False},
    ]

    df = app.m.json_to_dataframe(dados)

    assert str(df["Idade"].dtype) == "Int64" and int(df["Idade"].isna().sum()) == 1
    assert df["Nota"].dtype == "float64" and df["Ok"].tolist() == ["true", "false"]
    assert list(df.columns) == ["Nome", "Idade", "Nota", "Ok"]


def test_json_chaves_em_falta_e_aninhados(app):
    df = app.m.json_to_dataframe([{"a": 1}, {"b": {"x": [1, 2]}}])

    assert list(df.columns) == ["a", "b"] and pd.isna(df.loc[1, "a"]) and df.loc[1, "b"] == '{"x": [1, 2]}'


def test_json_misto_de_numeros_e_texto_vira_texto(app):
    df = app.m.json_to_dataframe([{"c": 1}, {"c": "x"}, {"c": 2.5}])

    assert df["c"].tolist() == ["1", "x", "2.5"]


def test_json_objecto_de_listas(app):
    df = app.m.json_to_dataframe({"a": [1, 2], "b": ["x", None]})

    assert df["a"].tolist() == [1, 2] and df["b"].tolist() == ["x", None]


@pytest.mark.parametrize("dados, mensagem", [
    ([], "não suportada"), ({}, "não suportada"), (5, "não suportada"), ("texto", "não suportada"),
    ([1, 2], "não suportada"), ({"a": 1}, "não suportada"),
    ({"a": [1], "b": [1, 2]}, "comprimentos"), ([{}], "não tem linhas"), ({"a": []}, "não tem linhas"),
])
def test_json_estruturas_invalidas(app, dados, mensagem):
    with pytest.raises(ValueError, match=mensagem):
        app.m.json_to_dataframe(dados)


def test_importar_json_pela_interface_com_datas_e_bom(app, tmp_path):
    caminho = tmp_path / "dados.json"
    caminho.write_text(json.dumps([{"D": "15/01/2020", "I": "2021-03-04", "N": 1}], ensure_ascii=False), encoding="utf-8-sig")

    app.run(app.m.import_json, str(caminho))

    assert app.m.col_types == {"D": app.m.DATE_TYPE, "I": app.m.DATE_TYPE, "N": int}
    assert app.m.table.loc[0, "I"] == pd.Timestamp("2021-03-04")


def test_importar_json_invalido_ou_inexistente_nao_altera_a_folha(app, tmp_path, capsys):
    app.dummy()
    mau = tmp_path / "mau.json"
    mau.write_text("{nao e json", encoding="utf-8")
    vazio = tmp_path / "vazio.json"
    vazio.write_text("[]", encoding="utf-8")

    app.run(app.m.import_json, "s", str(mau))
    app.run(app.m.import_json, "s", str(vazio))
    app.run(app.m.import_json, "s", str(tmp_path / "nada"))
    app.run(app.m.import_json, "n")  # recusa substituir

    saida = capsys.readouterr().out
    assert "JSON inválido" in saida and "não encontrado" in saida and len(app.m.table) == 8


def test_exportar_json_e_voltar_a_importar_sem_perdas(app, tmp_path):
    app.dummy()
    caminho = tmp_path / "saida.json"

    app.run(app.m.export_json, str(caminho))
    registos = json.loads(caminho.read_text(encoding="utf-8"))

    assert registos[0] == {"Nome": "João", "Idade": 25, "Salario": 1500.5, "Cidade": "Lisboa", "Admissao": "15/01/2020"}
    assert registos[3]["Idade"] is None  # em falta -> null
    original = app.m.table.copy()
    app.m.reset_workbook()
    app.run(app.m.import_json, str(caminho))
    assert app.m.table["Nome"].tolist() == original["Nome"].tolist()
    assert app.m.table["Admissao"].tolist() == original["Admissao"].tolist()
    assert app.m.col_types == {"Nome": str, "Idade": int, "Salario": float, "Cidade": str, "Admissao": app.m.DATE_TYPE}


# ---------------------------------------------------------------- texto colado


def test_colar_dados_do_excel_separados_por_tabulacao(app):
    app.run(app.m.import_pasted, "Nome\tIdade\tData", "Ana\t30\t01/02/2020", "Rui\t\t03/04/2021", "")

    assert list(app.m.table.columns) == ["Nome", "Idade", "Data"]
    assert app.m.col_types == {"Nome": str, "Idade": int, "Data": app.m.DATE_TYPE}
    assert int(app.m.table["Idade"].isna().sum()) == 1


def test_colar_csv_portugues_com_virgula_decimal_e_fim(app):
    app.run(app.m.import_pasted, "Produto;Preço", "Pão;1,5", "Leite;0,8", "fim")

    assert app.m.col_types["Preço"] is float and app.m.table["Preço"].tolist() == [1.5, 0.8]


def test_colar_para_com_linha_vazia_e_ignora_o_que_vem_depois(app):
    app.run(app.m.import_pasted, "A,B", "1,2", "", "linha que nunca é lida")

    assert len(app.m.table) == 1


def test_colar_cancelado_ou_vazio_nao_altera_nada(app, capsys):
    app.dummy()

    app.run(app.m.import_pasted, "s", "sair")
    app.run(app.m.import_pasted, "s", "")
    app.run(app.m.import_pasted, "n")

    assert len(app.m.table) == 8 and "não há dados" in capsys.readouterr().out


# ---------------------------------------------------------------- Markdown e HTML


def test_markdown_com_escape_alinhamento_e_formatos(app, tmp_path):
    app.m.table = pd.DataFrame({"Nome": ["a|b", "linha\nnova"], "V": [1234.5, 2.0]})
    app.m.col_types = {"Nome": str, "V": float}
    app.m.meta["formats"]["V"] = {"decimals": 1, "thousands": True, "locale": "intl"}
    caminho = tmp_path / "t.md"

    app.m.write_markdown(str(caminho))
    linhas = caminho.read_text(encoding="utf-8").splitlines()

    assert linhas[0] == "# Folha1" and "2 linhas × 2 colunas" in linhas[2]
    assert linhas[4] == "| Nome | V |" and linhas[5] == "| --- | ---: |"
    assert linhas[6] == "| a\\|b | 1,234.5 |" and linhas[7] == "| linha nova | 2.0 |"


def test_html_escapa_tudo_e_nao_tem_scripts(app, tmp_path):
    app.m.table = pd.DataFrame({"<b>col</b>": ['<script>alert("x")</script>', "R&D"], "N": [1, 2]})
    app.m.col_types = {"<b>col</b>": str, "N": int}
    app.m.apply_types()
    app.m.active_name = "A<i>"
    app.m.sheets = {"A<i>": None}
    caminho = tmp_path / "t.html"

    app.m.write_html(str(caminho))
    texto = caminho.read_text(encoding="utf-8")

    assert "<script" not in texto and "<b>col" not in texto and "&lt;script&gt;" in texto and "R&amp;D" in texto
    assert "<title>A&lt;i&gt;</title>" in texto and '<td class="num">1</td>' in texto and texto.startswith("<!doctype html>")


def test_exportar_markdown_e_html_pela_interface(app, tmp_path):
    app.dummy()

    app.run(app.m.export_markdown, str(tmp_path / "a"))
    app.run(app.m.export_html, str(tmp_path / "b.HTML"))

    assert (tmp_path / "a.md").exists() and (tmp_path / "b.HTML").exists()  # extensão sem distinguir maiúsculas


def test_exportar_sem_dados_avisa(app, capsys):
    app.m.export_json()

    assert "Sem dados" in capsys.readouterr().out


def test_exportar_com_erro_no_ficheiro_avisa_sem_rebentar(app, tmp_path, capsys):
    app.dummy()

    app.run(app.m.export_json, str(tmp_path / "pasta_que_nao_existe" / "x"))

    assert "Erro" in capsys.readouterr().out


# ---------------------------------------------------------------- projectos


def test_guardar_e_abrir_projecto_com_varias_folhas_e_metadados(app, tmp_path):
    app.dummy()
    app.m.meta["formats"]["Salario"] = {"decimals": 1}
    app.run(app.m.add_formula_column, "Dobro", "Idade * 2")
    app.m.add_sheet("Notas", pd.DataFrame({"T": ["x"]}), {"T": str}, activate=False)
    caminho = str(tmp_path / "meu")

    app.run(app.m.save_project, caminho)
    assert (tmp_path / "meu.termifolha").exists()

    app.m.reset_workbook()
    app.run(app.m.open_project, caminho)

    assert app.m.sheet_names() == ["Folha1", "Notas"] and app.m.active_name == "Folha1"
    assert app.m.meta["formulas"] == {"Dobro": "Idade * 2"} and app.m.meta["formats"] == {"Salario": {"decimals": 1}}
    app.m.switch_sheet("Notas")
    assert app.m.table["T"].tolist() == ["x"]


def test_abrir_projecto_invalido_ou_inexistente_nao_altera_o_livro(app, tmp_path, capsys):
    app.dummy()
    (tmp_path / "mau.termifolha").write_text('{"versao": 9}', encoding="utf-8")
    (tmp_path / "lixo.termifolha").write_text("nao e json", encoding="utf-8")

    app.run(app.m.open_project, "s", str(tmp_path / "mau"))
    app.run(app.m.open_project, "s", str(tmp_path / "lixo"))
    app.run(app.m.open_project, "s", str(tmp_path / "nada"))

    assert "Projecto inválido" in capsys.readouterr().out and len(app.m.table) == 8


def test_abrir_projecto_nao_importa_as_definicoes_do_ficheiro(app, tmp_path):
    app.dummy()
    app.m.settings["history_limit"] = 7
    app.run(app.m.save_project, str(tmp_path / "p"))
    app.m.settings["history_limit"] = 30

    app.run(app.m.open_project, "s", str(tmp_path / "p"))

    assert app.m.settings["history_limit"] == 30


# ---------------------------------------------------------------- definições


def test_reduzir_o_historico_corta_os_passos_antigos(app):
    app.dummy()
    for _ in range(5):
        app.run(app.m.remove_row, "0", "s")
    assert len(app.m.undo_stack) == 7

    app.m.set_history_limit(2)

    assert len(app.m.undo_stack) == 3 and app.m.settings["history_limit"] == 2


@pytest.mark.parametrize("valor", [0, 101, -1, True, "5", 2.5])
def test_limite_de_historico_invalido(app, valor):
    with pytest.raises(ValueError, match="entre 1 e"):
        app.m.set_history_limit(valor)


def test_menu_de_definicoes_altera_e_guarda(app):
    app.run(app.m.menu_settings, "1", "50", "2", "40", "0")

    assert app.m.settings == {"history_limit": 50, "page_size": 40}
    estado = json.loads(Path(app.m.STATE_FILE).read_text(encoding="utf-8"))
    assert estado["definicoes"] == {"history_limit": 50, "page_size": 40}


def test_menu_de_definicoes_valores_invalidos_mantem(app):
    app.run(app.m.menu_settings, "1", "500", "2", "abc", "0")

    assert app.m.settings == {"history_limit": app.m.HISTORY_LIMIT, "page_size": app.m.PAGE_SIZE}


def test_definicoes_sobrevivem_ao_reinicio(app):
    app.dummy()
    app.m.set_history_limit(9)
    app.m.save_state()

    app.m.settings.update({"history_limit": 20})
    app.m.reset_workbook()
    app.m.load_state()

    assert app.m.settings["history_limit"] == 9


def test_menu_principal_abre_definicoes_e_sai(app, capsys):
    app.run(app.m.run_interactive, "9", "0", "Q")

    assert "DEFINIÇÕES" in capsys.readouterr().out


# ---------------------------------------------------------------- modo de linha de comando


@pytest.fixture
def cli(app, monkeypatch):
    monkeypatch.setattr(app.m, "STATE_ENABLED", True)  # run_batch desliga-o; o monkeypatch repõe no fim
    monkeypatch.setattr(app.m, "configure_log", lambda: None)
    monkeypatch.setattr("builtins.input", lambda *_: pytest.fail("o modo não interactivo não pode perguntar nada"))
    return app


def _csv(tmp_path):
    caminho = tmp_path / "vendas.csv"
    caminho.write_text("Produto;Preco;Data\nPao;1,5;01/02/2020\nLeite;0,8;03/04/2021\n", encoding="utf-8")
    return str(caminho)


@pytest.mark.parametrize("extensao", [".xlsx", ".csv", ".json", ".md", ".html", ".docx", ".pdf"])
def test_cli_importa_csv_e_exporta_todos_os_formatos(cli, tmp_path, extensao, capsys):
    destino = tmp_path / f"saida{extensao}"

    codigo = cli.m.main(["--importar", _csv(tmp_path), "--exportar", str(destino)])

    assert codigo == 0 and destino.exists() and destino.stat().st_size > 0
    assert f"Exportado para '{destino}'" in capsys.readouterr().out


def test_cli_varias_exportacoes_e_csv_portugues(cli, tmp_path):
    codigo = cli.m.main(["--importar", _csv(tmp_path), "--exportar", str(tmp_path / "a.csv"), "--exportar", str(tmp_path / "b.json"), "--csv-pt"])

    assert codigo == 0
    assert (tmp_path / "a.csv").read_text(encoding="utf-8-sig").splitlines()[1] == "Pao;1,5;01/02/2020"
    assert json.loads((tmp_path / "b.json").read_text(encoding="utf-8"))[0]["Preco"] == 1.5


def test_cli_nao_le_nem_grava_o_estado_da_sessao(cli, tmp_path):
    cli.m.main(["--importar", _csv(tmp_path), "--exportar", str(tmp_path / "x.json")])

    assert not Path(cli.m.STATE_FILE).exists()


def test_cli_excel_com_todas_as_folhas_e_projecto(cli, tmp_path):
    cli.dummy()
    cli.m.add_sheet("Outra", pd.DataFrame({"X": [1]}), {"X": int})
    projecto = str(tmp_path / "p.termifolha")
    cli.m._write_state_file(projecto)
    cli.m.reset_workbook()

    codigo = cli.m.main(["--projecto", projecto, "--exportar", str(tmp_path / "livro.xlsx"), "--todas-as-folhas"])

    assert codigo == 0 and list(pd.read_excel(tmp_path / "livro.xlsx", sheet_name=None)) == ["Folha1", "Outra"]


def test_cli_importa_excel_com_varias_folhas_e_json(cli, tmp_path):
    livro = tmp_path / "l.xlsx"
    with pd.ExcelWriter(livro) as w:
        pd.DataFrame({"A": [1, 2]}).to_excel(w, sheet_name="Um", index=False)
        pd.DataFrame({"B": [3]}).to_excel(w, sheet_name="Dois", index=False)
    (tmp_path / "d.json").write_text('[{"k": 1}]', encoding="utf-8")

    assert cli.m.main(["--importar", str(livro), "--exportar", str(tmp_path / "o.json")]) == 0
    assert cli.m.sheet_names() == ["Um", "Dois"]
    assert cli.m.main(["--importar", str(tmp_path / "d.json"), "--exportar", str(tmp_path / "o2.csv")]) == 0


def test_cli_extensao_de_exportacao_desconhecida_falha_antes_de_importar(cli, tmp_path, capsys):
    codigo = cli.m.main(["--importar", _csv(tmp_path), "--exportar", str(tmp_path / "x.odt")])

    assert codigo == 2 and "não suportado" in capsys.readouterr().err and cli.m.table.columns.empty


@pytest.mark.parametrize("argumentos, mensagem", [
    (["--importar", "{tmp}/nao_existe.csv", "--exportar", "{tmp}/o.json"], "não encontrado"),
    (["--importar", "{tmp}/x.txt", "--exportar", "{tmp}/o.json"], "não suportado"),
    (["--exportar", "{tmp}/o.json"], "não há dados"),
    (["--projecto", "{tmp}/nao_existe.termifolha"], "não encontrado"),
])
def test_cli_erros_dao_codigo_1_e_mensagem_no_stderr(cli, tmp_path, capsys, argumentos, mensagem):
    (tmp_path / "x.txt").write_text("a", encoding="utf-8")

    codigo = cli.m.main([a.replace("{tmp}", str(tmp_path)) for a in argumentos])

    assert codigo == 1 and mensagem in capsys.readouterr().err


def test_cli_ficheiro_json_invalido_da_erro_sem_traceback(cli, tmp_path, capsys):
    mau = tmp_path / "mau.json"
    mau.write_text("{", encoding="utf-8")

    codigo = cli.m.main(["--importar", str(mau), "--exportar", str(tmp_path / "o.csv")])

    assert codigo == 1 and "Erro:" in capsys.readouterr().err


def test_cli_versao_e_argumentos_invalidos(cli, capsys):
    with pytest.raises(SystemExit) as saida:
        cli.m.main(["--versao"])
    assert saida.value.code == 0 and "TermiFolha 1.3" in capsys.readouterr().out

    with pytest.raises(SystemExit) as saida:
        cli.m.main(["--nao-existe"])
    assert saida.value.code == 2


def test_cli_a_serio_como_processo_separado(tmp_path):
    origem = tmp_path / "in.csv"
    origem.write_text("a,b\n1,x\n2,y\n", encoding="utf-8")
    destino = tmp_path / "out.json"

    ambiente = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    resultado = subprocess.run([sys.executable, str(RAIZ / "TermiFolha.py"), "--importar", str(origem), "--exportar", str(destino)],
                               cwd=tmp_path, capture_output=True, text=True, encoding="utf-8", env=ambiente, check=False)

    assert resultado.returncode == 0, resultado.stderr
    assert json.loads(destino.read_text(encoding="utf-8")) == [{"a": 1, "b": "x"}, {"a": 2, "b": "y"}]

    falha = subprocess.run([sys.executable, str(RAIZ / "TermiFolha.py"), "--importar", str(tmp_path / "nada.csv"), "--exportar", str(destino)],
                           cwd=tmp_path, capture_output=True, text=True, encoding="utf-8", env=ambiente, check=False)
    assert falha.returncode == 1 and "não encontrado" in falha.stderr
