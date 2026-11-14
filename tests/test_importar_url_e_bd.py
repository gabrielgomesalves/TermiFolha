"""Importar de endereços (URL) e de bases de dados (SQLite e, com o SQLAlchemy, outras). Sem rede: servidor HTTP local."""
import http.server
import socket
import sqlite3
import sys
import threading
import time

import pandas as pd
import pytest


@pytest.fixture(autouse=True)
def _repor_estado_activo(app, monkeypatch):
    """run_batch (usado por main) desliga a gravação do estado; o monkeypatch repõe-na no fim de cada teste."""
    monkeypatch.setattr(app.m, "STATE_ENABLED", True)


# ---------------------------------------------------------------- servidor HTTP local


class _Manipulador(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map, ".csv": "text/csv", ".json": "application/json"}

    def log_message(self, *args):
        pass

    def do_GET(self):  # noqa: N802 - nome imposto pela biblioteca
        if self.path.startswith("/lento"):
            time.sleep(1.0)
        super().do_GET()


@pytest.fixture
def servidor(tmp_path):
    """Serve a pasta `www` de tmp_path em http://127.0.0.1:<porta>/ ; devolve (pasta, endereço base)."""
    pasta = tmp_path / "www"
    pasta.mkdir()
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), lambda *a, **k: _Manipulador(*a, directory=str(pasta), **k))
    fio = threading.Thread(target=httpd.serve_forever, daemon=True)
    fio.start()
    yield pasta, f"http://127.0.0.1:{httpd.server_address[1]}"
    httpd.shutdown()
    httpd.server_close()


def _porta_livre():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


# ---------------------------------------------------------------- URL: utilitários

def test_is_url_e_safe_url(app):
    assert app.m.is_url("http://x.pt/a.csv") and app.m.is_url("  HTTPS://x.pt") and not app.m.is_url("ftp://x.pt") and not app.m.is_url("dados.csv")
    assert app.m.safe_url("https://user:pass@exemplo.pt:8443/a/b.csv?token=SEGREDO#frag") == "https://exemplo.pt:8443/a/b.csv"
    assert app.m.safe_url("http://exemplo.pt/x.json") == "http://exemplo.pt/x.json"


@pytest.mark.parametrize("endereco", ["ftp://x.pt/a.csv", "file:///etc/passwd", "dados.csv", "javascript:alert(1)"])
def test_so_aceita_http_e_https(app, endereco):
    with pytest.raises(ValueError, match="http:// ou https://"):
        app.m.fetch_url(endereco)


# ---------------------------------------------------------------- URL: importação

def test_importar_csv_portugues_por_url(app, servidor):
    pasta, base = servidor
    (pasta / "vendas.csv").write_text("Nome;Valor;Data\nAna;1,5;15/01/2020\nRui;2,25;03/02/2021\n", encoding="utf-8")

    app.run(app.m.import_url, f"{base}/vendas.csv")

    assert list(app.m.table.columns) == ["Nome", "Valor", "Data"] and len(app.m.table) == 2
    assert app.m.col_types == {"Nome": str, "Valor": float, "Data": app.m.DATE_TYPE}
    assert app.m.table.loc[1, "Valor"] == 2.25


def test_csv_em_latin1_por_url(app, servidor):
    pasta, base = servidor
    (pasta / "acentos.csv").write_bytes("Cidade,N\nÉvora,1\nSão João,2\n".encode("latin-1"))

    app.run(app.m.import_url, f"{base}/acentos.csv")

    assert app.m.table["Cidade"].tolist() == ["Évora", "São João"]


def test_importar_json_por_url(app, servidor):
    pasta, base = servidor
    (pasta / "dados.json").write_text('[{"nome": "Ana", "idade": 30}, {"nome": "Rui", "idade": null}]', encoding="utf-8")

    app.run(app.m.import_url, f"{base}/dados.json")

    assert app.m.table["nome"].tolist() == ["Ana", "Rui"] and app.m.col_types["idade"] is int


def test_importar_excel_com_varias_folhas_por_url(app, servidor):
    pasta, base = servidor
    with pd.ExcelWriter(pasta / "livro.xlsx") as escritor:
        pd.DataFrame({"a": [1, 2]}).to_excel(escritor, sheet_name="Primeira", index=False)
        pd.DataFrame({"b": ["x", "y", "z"]}).to_excel(escritor, sheet_name="Segunda", index=False)

    app.run(app.m.import_url, f"{base}/livro.xlsx")

    assert app.m.sheet_names() == ["Primeira", "Segunda"] and len(app.m.get_sheet("Segunda")["table"]) == 3


def test_excel_de_uma_folha_por_url_substitui_a_activa(app, servidor):
    pasta, base = servidor
    pd.DataFrame({"a": [1, 2, 3]}).to_excel(pasta / "uma.xlsx", index=False)

    app.run(app.m.import_url, f"{base}/uma.xlsx")

    assert app.m.sheet_names() == ["Folha1"] and app.m.table["a"].tolist() == [1, 2, 3]


def test_formato_adivinhado_pelo_conteudo_quando_nao_ha_extensao(app, servidor):
    pasta, base = servidor
    (pasta / "api").write_text('{"x": [1, 2, 3], "y": ["a", "b", "c"]}', encoding="utf-8")
    (pasta / "tabela").write_text("a,b\n1,2\n3,4\n", encoding="utf-8")
    pd.DataFrame({"q": [7, 8]}).to_excel(pasta / "livro.dat", index=False, engine="openpyxl")

    app.run(app.m.import_url, f"{base}/api")
    assert app.m.table["y"].tolist() == ["a", "b", "c"]

    app.run(app.m.import_url, "s", f"{base}/tabela")  # já há folha: confirma a substituição
    assert app.m.table["b"].tolist() == [2, 4]

    app.run(app.m.import_url, "s", f"{base}/livro.dat")  # Excel reconhecido pelo início do ficheiro (PK)
    assert app.m.table["q"].tolist() == [7, 8]


def test_pede_confirmacao_antes_de_substituir_e_nao_descarrega_se_recusar(app, servidor, monkeypatch):
    _, base = servidor
    app.dummy()
    monkeypatch.setattr(app.m, "fetch_url", lambda *a, **k: pytest.fail("não devia descarregar"))

    app.run(app.m.import_url, "n")

    assert len(app.m.table) == 8


@pytest.mark.parametrize("resposta", ["", "sair"])
def test_cancelar_nao_faz_nada(app, resposta):
    app.run(app.m.import_url, resposta)

    assert app.m.table.empty


def test_erro_404_nao_mostra_a_chave_do_endereco(app, servidor, capsys):
    _, base = servidor

    app.run(app.m.import_url, f"{base}/nao_existe.csv?token=SEGREDO")

    saida = capsys.readouterr().out
    assert "o servidor respondeu 404" in saida and "SEGREDO" not in saida and "nao_existe.csv" in saida


def test_servidor_inalcancavel(app, capsys):
    app.run(app.m.import_url, f"http://127.0.0.1:{_porta_livre()}/x.csv?chave=SEGREDO")

    saida = capsys.readouterr().out
    assert "não foi possível ligar" in saida and "SEGREDO" not in saida


def test_tempo_limite(app, servidor):
    _, base = servidor

    with pytest.raises(ValueError, match="não respondeu"):
        app.m.fetch_url(f"{base}/lento", timeout=0.2)


def test_ficheiro_grande_demais(app, servidor, monkeypatch):
    pasta, base = servidor
    (pasta / "grande.csv").write_bytes(b"a\n" + b"1\n" * 600_000)  # ~1,2 MB
    monkeypatch.setattr(app.m, "URL_MAX_BYTES", 1024 * 1024)

    with pytest.raises(ValueError, match="mais de 1 MB"):
        app.m.read_url_frames(f"{base}/grande.csv")


def test_conteudo_invalido_diz_qual_era_o_formato(app, servidor):
    pasta, base = servidor
    (pasta / "mau.xlsx").write_text("isto não é um Excel", encoding="utf-8")
    (pasta / "mau.json").write_text("{isto não é json", encoding="utf-8")
    (pasta / "vazio.csv").write_text("", encoding="utf-8")

    with pytest.raises(ValueError, match="não é um xlsx válido"):
        app.m.read_url_frames(f"{base}/mau.xlsx")
    with pytest.raises(ValueError, match="não é um json válido"):
        app.m.read_url_frames(f"{base}/mau.json")
    with pytest.raises(ValueError, match="não é um csv válido"):
        app.m.read_url_frames(f"{base}/vazio.csv")


def test_xls_antigo_e_recusado(app, servidor):
    pasta, base = servidor
    (pasta / "velho.xls").write_bytes(b"\xd0\xcf\x11\xe0" + b"\x00" * 100)

    with pytest.raises(ValueError, match=r"\.xls antigo"):
        app.m.read_url_frames(f"{base}/velho.xls")


# ---------------------------------------------------------------- URL: linha de comandos

def test_cli_importa_de_url_e_exporta(app, servidor, tmp_path):
    pasta, base = servidor
    (pasta / "d.csv").write_text("a,b\n1,x\n2,y\n", encoding="utf-8")
    saida = tmp_path / "saida.json"

    codigo = app.m.main(["--importar", f"{base}/d.csv", "--exportar", str(saida)])

    assert codigo == 0 and '"a": 1' in saida.read_text(encoding="utf-8")


def test_cli_url_inalcancavel_devolve_1(app, tmp_path, capsys):
    codigo = app.m.main(["--importar", f"http://127.0.0.1:{_porta_livre()}/x.csv", "--exportar", str(tmp_path / "s.csv")])

    assert codigo == 1 and "não foi possível ligar" in capsys.readouterr().err


# ---------------------------------------------------------------- bases de dados (SQLite)

@pytest.fixture
def bd(tmp_path):
    """Base SQLite com clientes, pedidos, uma vista, um binário e nomes de tabela difíceis. Devolve o caminho."""
    caminho = tmp_path / "loja.db"
    ligacao = sqlite3.connect(caminho)
    ligacao.executescript("""
        CREATE TABLE clientes (id INTEGER PRIMARY KEY, nome TEXT, nascimento TEXT, saldo REAL, nota INTEGER);
        INSERT INTO clientes VALUES (1, 'Ana', '15/01/1990', 10.5, 5), (2, 'Rui', '03/02/1985', 20.0, NULL), (3, 'Eva', '07/07/2001', 0.25, 3);
        CREATE TABLE pedidos (id INTEGER PRIMARY KEY, cliente_id INTEGER, valor REAL);
        INSERT INTO pedidos VALUES (1, 1, 100.0), (2, 1, 50.0), (3, 3, 7.5);
        CREATE VIEW totais AS SELECT cliente_id, SUM(valor) AS total FROM pedidos GROUP BY cliente_id;
        CREATE TABLE "meu quadro" (a INTEGER, b BLOB);
        INSERT INTO "meu quadro" VALUES (1, x'0102030405');
        CREATE TABLE "x:y" (v INTEGER);
        INSERT INTO "x:y" VALUES (1);
        CREATE TABLE "x_y" (v INTEGER);
        INSERT INTO "x_y" VALUES (2);
        CREATE TABLE vazia (c INTEGER);
    """)
    ligacao.commit()
    ligacao.close()
    return caminho


def test_listar_tabelas_e_vistas_sem_as_internas(app, bd):
    assert app.m.list_tables(str(bd)) == ["clientes", "meu quadro", "pedidos", "totais", "vazia", "x:y", "x_y"]


def test_ler_tabela_sem_distinguir_maiusculas_e_com_nomes_dificeis(app, bd):
    assert len(app.m.read_table(str(bd), "CLIENTES")) == 3
    quadro = app.m.read_table(str(bd), "meu quadro")
    assert quadro["b"].tolist() == ["<binário, 5 bytes>"]


def test_tabela_inexistente_lista_as_que_existem(app, bd):
    with pytest.raises(ValueError, match=r"a tabela 'nada' não existe \(tabelas: clientes"):
        app.m.read_table(str(bd), "nada")


def test_consulta_com_juncao(app, bd):
    df = app.m.read_query(str(bd), "SELECT c.nome, SUM(p.valor) AS total FROM clientes c JOIN pedidos p ON p.cliente_id = c.id GROUP BY c.nome ORDER BY c.nome;")

    assert df.to_dict("list") == {"nome": ["Ana", "Eva"], "total": [150.0, 7.5]}


def test_consulta_com_with_e_funcao_replace_sao_permitidas(app, bd):
    df = app.m.read_query(str(bd), "WITH t AS (SELECT nome FROM clientes) SELECT replace(nome, 'a', 'A') AS n FROM t ORDER BY nome")

    assert df["n"].tolist() == ["AnA", "EvA", "Rui"]


@pytest.mark.parametrize("consulta, erro", [
    ("", "está vazia"),
    ("   ;  ", "está vazia"),
    ("DELETE FROM clientes", "começar por SELECT"),
    ("-- comentário\nSELECT 1", "começar por SELECT"),
    ("SELECT 1; DROP TABLE clientes", "uma consulta de cada vez"),
    ("SELECT * FROM clientes WHERE nome = 'update'", "só pode ler dados"),
    ("WITH t AS (SELECT 1) INSERT INTO clientes SELECT * FROM t", "só pode ler dados"),
    ("select * from clientes; pragma writable_schema=1", "uma consulta de cada vez"),
])
def test_consultas_que_nao_leem_sao_recusadas(app, bd, consulta, erro):
    with pytest.raises(ValueError, match=erro):
        app.m.read_query(str(bd), consulta)

    assert len(app.m.read_table(str(bd), "clientes")) == 3  # nada foi alterado


def test_a_ligacao_e_so_de_leitura_mesmo_que_a_verificacao_falhasse(app, bd):
    with app.m._open_database(str(bd)) as ligacao, pytest.raises(sqlite3.OperationalError):
        ligacao.execute("INSERT INTO clientes VALUES (9, 'X', '', 0, 0)")

    assert len(app.m.read_table(str(bd), "clientes")) == 3


def test_ficheiro_inexistente_e_ficheiro_que_nao_e_base_de_dados(app, tmp_path):
    with pytest.raises(FileNotFoundError):
        app.m.list_tables(str(tmp_path / "nao_existe.db"))
    falso = tmp_path / "falso.db"
    falso.write_text("isto não é SQLite" * 50, encoding="utf-8")

    with pytest.raises(ValueError, match="erro na base de dados"):
        app.m.list_tables(str(falso))


def test_consulta_com_erro_de_sql_da_mensagem_da_base_de_dados(app, bd):
    with pytest.raises(ValueError, match="erro na base de dados.*no such (table|column)"):
        app.m.read_query(str(bd), "SELECT x FROM nao_existe")


def test_consulta_sem_linhas(app, bd):
    with pytest.raises(ValueError, match="não devolveu linhas"):
        app.m.read_query(str(bd), "SELECT * FROM clientes WHERE id > 100")
    with pytest.raises(ValueError, match="não devolveu linhas"):
        app.m.read_table(str(bd), "vazia")


def test_limite_de_linhas(app, bd, monkeypatch):
    monkeypatch.setattr(app.m, "DB_MAX_ROWS", 2)
    monkeypatch.setattr(app.m, "DB_CHUNK_ROWS", 1)

    with pytest.raises(ValueError, match="mais de 2 linhas"):
        app.m.read_table(str(bd), "clientes")


def test_colunas_repetidas_na_consulta_ficam_unicas(app, bd):
    df = app.m.read_query(str(bd), "SELECT c.id, p.id, p.valor FROM clientes c JOIN pedidos p ON p.cliente_id = c.id")

    assert list(df.columns) == ["id", "id_2", "valor"]


def test_todas_as_tabelas_falha_se_alguma_estiver_vazia(app, bd):
    with pytest.raises(ValueError, match="não devolveu linhas"):  # a tabela 'vazia' não pode virar folha
        app.m.read_all_tables(str(bd))


def test_endereco_sqlite_com_barras(app, bd):
    endereco = "sqlite:///" + bd.as_posix()

    assert app.m.list_tables(endereco)[0] == "clientes" and len(app.m.read_table(endereco, "pedidos")) == 3


def test_valid_sheet_names_unicos_e_validos(app):
    nomes = app.m._valid_sheet_names(["x:y", "x_y", "a" * 50, "[ab]", "", "aspas'e\"", "X_Y"])

    assert nomes[0] == "x_y" and nomes[1] == "x_y (2)" and nomes[6] == "X_Y (3)"
    assert len(nomes[2]) == 31 and nomes[3] == "_ab_" and nomes[4] == "Folha" and nomes[5] == "aspas'e_"
    assert all(app.m.sheet_name_error(n, nomes[:i]) is None for i, n in enumerate(nomes))


# ---------------------------------------------------------------- bases de dados: interface

def test_importar_tabela_por_nome_com_datas_e_inteiros(app, bd):
    # [coluna 'nota' tem um NULL: o pandas lê-a como decimal, por isso o programa pergunta]
    app.run(app.m.import_database, str(bd), "clientes", "s")

    t = app.m.table
    assert list(t.columns) == ["id", "nome", "nascimento", "saldo", "nota"]
    assert app.m.col_types["nascimento"] is app.m.DATE_TYPE and app.m.col_types["nota"] is int
    assert app.m.col_types["saldo"] is float and t["nota"].isna().sum() == 1


def test_importar_tabela_por_numero_e_por_consulta(app, bd):
    app.run(app.m.import_database, str(bd), "3")  # [3] = pedidos
    assert list(app.m.table.columns) == ["id", "cliente_id", "valor"] and len(app.m.table) == 3

    app.run(app.m.import_database, "s", str(bd), "SELECT nome FROM clientes ORDER BY nome DESC LIMIT 2")
    assert app.m.table["nome"].tolist() == ["Rui", "Eva"]


def test_importar_todas_as_tabelas_exceptuando_as_vazias_falha_com_mensagem(app, bd, capsys):
    app.run(app.m.import_database, str(bd), "*")

    assert "não devolveu linhas" in capsys.readouterr().out


def test_importar_todas_as_tabelas(app, tmp_path):
    caminho = tmp_path / "pequena.db"
    ligacao = sqlite3.connect(caminho)
    ligacao.executescript('CREATE TABLE "x:y" (v INTEGER); INSERT INTO "x:y" VALUES (1); CREATE TABLE x_y (v INTEGER); INSERT INTO x_y VALUES (2), (3);')
    ligacao.commit()
    ligacao.close()

    app.run(app.m.import_database, str(caminho), "*")

    assert app.m.sheet_names() == ["x_y", "x_y (2)"] and len(app.m.get_sheet("x_y (2)")["table"]) == 2


@pytest.mark.parametrize("resposta", ["", "sair"])
def test_cancelar_a_importacao_de_bd(app, bd, resposta):
    app.run(app.m.import_database, resposta)
    app.run(app.m.import_database, str(bd), resposta)

    assert app.m.table.empty


def test_mensagens_de_erro_da_interface_de_bd(app, bd, tmp_path, capsys):
    app.run(app.m.import_database, str(tmp_path / "nada.db"))
    assert "não encontrado" in capsys.readouterr().out

    app.run(app.m.import_database, str(bd), "9")
    assert "a tabela '9' não existe" in capsys.readouterr().out

    app.run(app.m.import_database, str(bd), "DROP TABLE clientes")
    assert "a tabela 'DROP TABLE clientes' não existe" in capsys.readouterr().out

    app.run(app.m.import_database, str(bd), "SELECT 1; DELETE FROM clientes")
    assert "uma consulta de cada vez" in capsys.readouterr().out


def test_bd_sem_tabelas(app, tmp_path, capsys):
    vazia = tmp_path / "vazia.db"
    sqlite3.connect(vazia).close()

    app.run(app.m.import_database, str(vazia))

    assert "não tem tabelas" in capsys.readouterr().out


def test_pede_confirmacao_antes_de_substituir_na_bd(app, bd):
    app.dummy()

    app.run(app.m.import_database, "n")

    assert len(app.m.table) == 8


# ---------------------------------------------------------------- bases de dados: linha de comandos

def test_cli_tabela_consulta_e_todas(app, bd, tmp_path):
    csv = tmp_path / "t.csv"
    assert app.m.main(["--base-dados", str(bd), "--tabela", "pedidos", "--exportar", str(csv)]) == 0
    assert len(pd.read_csv(csv)) == 3

    assert app.m.main(["--base-dados", str(bd), "--consulta", "SELECT nome FROM clientes WHERE id = 1", "--exportar", str(csv)]) == 0
    assert pd.read_csv(csv)["nome"].tolist() == ["Ana"]


def test_cli_todas_as_tabelas_num_excel(app, tmp_path):
    caminho = tmp_path / "duas.db"
    ligacao = sqlite3.connect(caminho)
    ligacao.executescript("CREATE TABLE a (v INTEGER); INSERT INTO a VALUES (1); CREATE TABLE b (w TEXT); INSERT INTO b VALUES ('x'), ('y');")
    ligacao.commit()
    ligacao.close()
    saida = tmp_path / "livro.xlsx"

    codigo = app.m.main(["--base-dados", str(caminho), "--tabela", "*", "--exportar", str(saida), "--todas-as-folhas"])

    assert codigo == 0 and list(pd.read_excel(saida, sheet_name=None)) == ["a", "b"]


def test_cli_argumentos_invalidos_devolvem_2(app, bd, tmp_path, capsys):
    saida = str(tmp_path / "s.csv")

    assert app.m.main(["--base-dados", str(bd), "--exportar", saida]) == 2
    assert "indique --tabela" in capsys.readouterr().err
    assert app.m.main(["--tabela", "x", "--exportar", saida]) == 2
    assert "só se usam com --base-dados" in capsys.readouterr().err
    with pytest.raises(SystemExit) as e:
        app.m.main(["--base-dados", str(bd), "--tabela", "a", "--consulta", "SELECT 1"])
    assert e.value.code == 2
    with pytest.raises(SystemExit) as e:
        app.m.main(["--importar", "a.csv", "--base-dados", str(bd), "--tabela", "a"])
    assert e.value.code == 2


def test_cli_erros_de_bd_devolvem_1_com_mensagem(app, bd, tmp_path, capsys):
    saida = str(tmp_path / "s.csv")

    assert app.m.main(["--base-dados", str(bd), "--tabela", "nada", "--exportar", saida]) == 1
    assert "a tabela 'nada' não existe" in capsys.readouterr().err
    assert app.m.main(["--base-dados", str(tmp_path / "x.db"), "--tabela", "a", "--exportar", saida]) == 1
    assert "não encontrado" in capsys.readouterr().err
    assert app.m.main(["--base-dados", str(bd), "--consulta", "DROP TABLE clientes", "--exportar", saida]) == 1
    assert len(app.m.read_table(str(bd), "clientes")) == 3


# ---------------------------------------------------------------- outras bases de dados (SQLAlchemy opcional)

def test_sem_sqlalchemy_explica_o_que_instalar(app, monkeypatch):
    monkeypatch.setitem(sys.modules, "sqlalchemy", None)  # simula o pacote não instalado

    with pytest.raises(ValueError, match="instale o SQLAlchemy"):
        app.m.list_tables("postgresql://utilizador:segredo@servidor/base")


def test_palavra_passe_nunca_aparece_nas_mensagens(app):
    texto = "falha ao ligar a postgresql://utilizador:segredo123@servidor/base"

    assert "segredo123" not in app.m._mask_secrets(texto, "postgresql://utilizador:segredo123@servidor/base")
    assert app.m._mask_secrets(texto, "sqlite:///x.db") == texto
    assert app.m._mask_secrets("ok", "http://[mal") == "ok"


def test_sqlalchemy_le_tabelas_e_consultas(app, bd):
    pytest.importorskip("sqlalchemy")
    endereco = "sqlite+pysqlite:///" + bd.as_posix()  # um dialecto que não seja o 'sqlite' simples passa pelo SQLAlchemy

    assert app.m._sqlite_path(endereco) is None
    assert app.m.list_tables(endereco) == ["clientes", "meu quadro", "pedidos", "totais", "vazia", "x:y", "x_y"]
    assert app.m.read_table(endereco, "clientes")["nome"].tolist() == ["Ana", "Rui", "Eva"]
    assert app.m.read_table(endereco, "meu quadro")["b"].tolist() == ["<binário, 5 bytes>"]
    assert app.m.read_query(endereco, "SELECT SUM(valor) AS s FROM pedidos")["s"].tolist() == [157.5]
    with pytest.raises(ValueError, match="erro na base de dados"):
        app.m.read_query(endereco, "SELECT x FROM nao_existe")
    with pytest.raises(ValueError, match="a tabela 'nada' não existe"):
        app.m.read_table(endereco, "nada")


def test_sqlalchemy_endereco_invalido_ou_sem_driver(app):
    pytest.importorskip("sqlalchemy")

    with pytest.raises(ValueError, match="inválido ou driver em falta"):
        app.m.list_tables("nao_existe_dialecto://utilizador:segredo123@servidor/base")


def test_caminho_da_bd(app):
    assert app.m._sqlite_path("dados.db") == "dados.db"
    assert app.m._sqlite_path("sqlite:///rel/dados.db") == "rel/dados.db"
    assert app.m._sqlite_path("sqlite:////abs/dados%20x.db") == "/abs/dados x.db"
    assert app.m._sqlite_path("postgresql://x/y") is None
