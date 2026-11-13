import logging
import subprocess
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent


def test_importar_o_modulo_nao_configura_ficheiro_de_log():
    codigo = (
        f"import sys, logging; sys.path.insert(0, {str(RAIZ)!r}); import TermiFolha as T; "
        "print([type(h).__name__ for h in logging.getLogger('termifolha').handlers])"
    )
    saida = subprocess.run([sys.executable, "-c", codigo], capture_output=True, text=True, check=True).stdout.strip()

    assert saida == "['NullHandler']"


def test_falha_de_importacao_fica_registada_com_traceback(app, tmp_path, caplog):
    (tmp_path / "vazio.csv").write_text("", encoding="utf-8")

    with caplog.at_level(logging.ERROR, logger="termifolha"):
        app.run(app.m.importar_csv, "vazio")

    registo = next(r for r in caplog.records if "importar_csv" in r.getMessage())
    assert registo.exc_info is not None


def test_falha_a_gravar_estado_fica_registada(app, tmp_path, caplog):
    app.m.FICHEIRO_ESTADO = str(tmp_path / "inexistente" / "e.json")

    with caplog.at_level(logging.ERROR, logger="termifolha"):
        app.m.guardar_estado()

    assert any("guardar o estado" in r.getMessage() for r in caplog.records)


def test_configurar_log_escreve_no_ficheiro(app, tmp_path):
    ficheiro = tmp_path / "t.log"
    app.m.FICHEIRO_LOG = str(ficheiro)
    antes = list(app.m.log.handlers)
    try:
        app.m.configurar_log()
        app.m.log.error("mensagem de teste")
        for h in app.m.log.handlers:
            h.flush()
        assert "mensagem de teste" in ficheiro.read_text(encoding="utf-8")
    finally:
        for h in list(app.m.log.handlers):
            if h not in antes:
                app.m.log.removeHandler(h)
                h.close()
        app.m.log.setLevel(logging.NOTSET)


def test_configurar_log_sem_permissoes_nao_rebenta(app, tmp_path):
    app.m.FICHEIRO_LOG = str(tmp_path / "nao_existe" / "t.log")
    antes = len(app.m.log.handlers)

    app.m.configurar_log()

    assert len(app.m.log.handlers) == antes


@pytest.mark.parametrize("operacao", ["exportar_pdf", "exportar_word", "exportar_csv", "exportar"])
def test_falha_de_exportacao_fica_registada(app, tmp_path, caplog, operacao):
    app.dummy()

    with caplog.at_level(logging.ERROR, logger="termifolha"):
        app.run(getattr(app.m, operacao), "pasta_que_nao_existe/saida")

    assert any(operacao in r.getMessage() for r in caplog.records)
