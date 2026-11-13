import os
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent


def test_ficheiro_de_estado_fica_ao_lado_do_programa(tmp_path):
    """Regressão M3: o caminho era relativo ao directório actual, não ao programa."""
    saida = subprocess.run(
        [sys.executable, "-c", f"import sys; sys.path.insert(0, {str(RAIZ)!r}); import TermiFolha; print(TermiFolha.FICHEIRO_ESTADO)"],
        cwd=tmp_path, capture_output=True, text=True, check=True,
    ).stdout.strip()

    assert os.path.isabs(saida)
    assert Path(saida).parent == RAIZ


def test_estado_corrompido_e_preservado_e_utilizador_e_avisado(app, capsys):
    """Regressão M1: o erro era engolido e o ficheiro depois sobrescrito."""
    caminho = Path(app.m.FICHEIRO_ESTADO)
    caminho.write_bytes(b"isto nao e um pickle")

    app.m.carregar_estado()

    saida = capsys.readouterr().out
    assert "corromp" in saida.lower()
    assert app.m.tabela.empty
    assert not caminho.exists()
    assert Path(str(caminho) + ".corrompido").read_bytes() == b"isto nao e um pickle"


def test_falha_a_gravar_estado_avisa_o_utilizador(app, capsys, tmp_path):
    """Regressão M1: falhas de gravação passavam despercebidas."""
    app.m.FICHEIRO_ESTADO = str(tmp_path)  # um directório: open(..., 'wb') falha

    app.m.guardar_estado()

    assert "não foi possível guardar" in capsys.readouterr().out.lower()


def test_estado_guardado_e_recarregado(app):
    app.dummy()
    linhas = len(app.m.tabela)

    app.m.tabela = app.m.pd.DataFrame()
    app.m.tipos = {}
    app.m.carregar_estado()

    assert len(app.m.tabela) == linhas
    assert app.m.tipos["Idade"] is int
