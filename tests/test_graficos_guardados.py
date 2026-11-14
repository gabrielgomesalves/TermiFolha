"""Título e rótulos dos gráficos: ficam guardados com a folha (por gráfico ou para todos) e sobrevivem ao reiniciar."""
import json
from pathlib import Path


def _capturar(app, monkeypatch):
    """Intercepta plt.show e devolve a lista de (título, rótulo X, rótulo Y) de cada gráfico mostrado."""
    vistos = []
    monkeypatch.setattr(app.m.plt, "show", lambda: vistos.append([(a.get_title(), a.get_xlabel(), a.get_ylabel()) for a in app.m.plt.gcf().axes]))
    return vistos


def test_textos_valem_para_todos_os_graficos_e_nao_se_gastam(app, monkeypatch):
    app.dummy()
    vistos = _capturar(app, monkeypatch)

    app.run(app.m.set_plot_labels, "0", "Vendas 2026", "Pessoa", "Euros")
    app.run(app.m.plot_scatter, "Idade,Salario")
    app.run(app.m.plot_scatter, "Idade,Salario")  # continua a valer (antes só valia uma vez)
    app.run(app.m.plot_area, "Salario")

    assert vistos[0] == vistos[1] == [("Vendas 2026", "Pessoa", "Euros")]
    assert vistos[2][0][0] == "Vendas 2026"


def test_texto_de_um_grafico_sobrepoe_se_ao_de_todos(app, monkeypatch):
    app.dummy()
    vistos = _capturar(app, monkeypatch)

    app.run(app.m.set_plot_labels, "0", "Geral", "", "")
    app.run(app.m.set_plot_labels, "4", "Só dispersão", "", "")  # [4] = Dispersão
    app.run(app.m.plot_scatter, "Idade,Salario")
    app.run(app.m.plot_area, "Salario")

    assert vistos[0][0] == ("Só dispersão", "Idade", "Salario")  # título do gráfico; eixos automáticos
    assert vistos[1][0][0] == "Geral"


def test_so_os_campos_preenchidos_mudam(app, monkeypatch):
    app.dummy()
    vistos = _capturar(app, monkeypatch)

    app.run(app.m.set_plot_labels, "0", "", "", "Y")
    app.run(app.m.plot_area, "Salario")

    assert vistos[0][0][0] == "Evolução da soma acumulada" and vistos[0][0][2] == "Y"


def test_enter_mantem_e_hifen_apaga(app):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "T", "X", "Y")

    app.run(app.m.set_plot_labels, "0", "", "-", "")  # mantém o título e o eixo Y, apaga o eixo X

    assert app.m.meta["charts"] == {"*": {"title": "T", "ylabel": "Y"}}


def test_apagar_todos_os_campos_remove_a_entrada(app):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "T", "", "")

    app.run(app.m.set_plot_labels, "0", "-", "", "")

    assert app.m.meta["charts"] == {}


def test_apagar_tudo_pede_confirmacao(app):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "T", "", "")

    app.run(app.m.set_plot_labels, "L", "n")
    assert app.m.meta["charts"] == {"*": {"title": "T"}}

    app.run(app.m.set_plot_labels, "L", "s")
    assert app.m.meta["charts"] == {}


def test_apagar_tudo_sem_textos_nao_pergunta_nada(app, capsys):
    app.dummy()

    app.run(app.m.set_plot_labels, "L")

    assert "Não há textos guardados" in capsys.readouterr().out


def test_opcao_invalida_e_cancelar(app, capsys):
    app.dummy()

    app.run(app.m.set_plot_labels, "99")
    assert "Opção inválida" in capsys.readouterr().out
    app.run(app.m.set_plot_labels, "sair")
    app.run(app.m.set_plot_labels, "")

    assert app.m.meta["charts"] == {}


def test_cada_folha_tem_os_seus_textos(app, monkeypatch):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "Folha um", "", "")
    app.m.add_sheet("Outra")
    app.run(app.m.create_dummy_sheet)
    vistos = _capturar(app, monkeypatch)

    app.run(app.m.plot_area, "Salario")
    assert vistos[0][0][0] == "Evolução da soma acumulada"

    app.m.switch_sheet("Folha1")
    app.run(app.m.plot_area, "Salario")
    assert vistos[1][0][0] == "Folha um"


def test_textos_ficam_gravados_no_estado_e_voltam_ao_reiniciar(app, monkeypatch):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "Vendas", "", "Euros")
    app.run(app.m.set_plot_labels, "1", "Barras do ano", "", "")

    estado = json.loads(Path(app.m.STATE_FILE).read_text(encoding="utf-8"))
    assert estado["folhas"][0]["graficos"] == {"*": {"title": "Vendas", "ylabel": "Euros"}, "bars": {"title": "Barras do ano"}}

    app.m.reset_workbook()
    assert app.m.meta["charts"] == {}
    app.m.load_state()
    assert app.m.meta["charts"] == {"*": {"title": "Vendas", "ylabel": "Euros"}, "bars": {"title": "Barras do ano"}}


def test_estado_sem_a_chave_dos_graficos_continua_a_abrir(app):
    app.dummy()
    estado = json.loads(Path(app.m.STATE_FILE).read_text(encoding="utf-8"))
    del estado["folhas"][0]["graficos"]
    Path(app.m.STATE_FILE).write_text(json.dumps(estado), encoding="utf-8")

    app.m.reset_workbook()
    app.m.load_state()

    assert app.m.meta["charts"] == {} and len(app.m.table) == 8


def test_estado_adulterado_nao_deixa_passar_textos_invalidos(app):
    app.dummy()
    estado = json.loads(Path(app.m.STATE_FILE).read_text(encoding="utf-8"))
    estado["folhas"][0]["graficos"] = {
        "*": {"title": "  bom  $x$ ", "xlabel": 5, "outro": "ignorado", "ylabel": "   "},
        "inventado": {"title": "x"},
        "bars": "texto",
        "pie": {"title": "y" * 500},
    }
    Path(app.m.STATE_FILE).write_text(json.dumps(estado), encoding="utf-8")

    app.m.reset_workbook()
    app.m.load_state()

    assert app.m.meta["charts"]["*"] == {"title": r"bom \$x\$"}
    assert set(app.m.meta["charts"]) == {"*", "pie"}
    assert len(app.m.meta["charts"]["pie"]["title"]) == app.m.PLOT_TEXT_MAX


def test_estado_com_graficos_que_nao_e_dicionario_e_ignorado(app):
    app.dummy()
    estado = json.loads(Path(app.m.STATE_FILE).read_text(encoding="utf-8"))
    estado["folhas"][0]["graficos"] = ["x"]
    Path(app.m.STATE_FILE).write_text(json.dumps(estado), encoding="utf-8")

    app.m.reset_workbook()
    app.m.load_state()

    assert app.m.meta["charts"] == {}


def test_desfazer_repoe_os_textos_anteriores(app):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "Primeiro", "", "")
    app.run(app.m.set_plot_labels, "0", "Segundo", "", "")

    app.run(app.m.undo_last)

    assert app.m.meta["charts"] == {"*": {"title": "Primeiro"}}


def test_duplicar_folha_copia_os_textos(app):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "Título", "", "")

    app.m.duplicate_sheet("Folha1", "Cópia")
    app.m.meta["charts"]["*"]["title"] = "Mudado"

    assert app.m.get_sheet("Folha1")["meta"]["charts"]["*"]["title"] == "Título"


def test_textos_tambem_valem_ao_gravar_png(app, tmp_path, monkeypatch):
    app.dummy()
    monkeypatch.setattr(app.m, "plot_png_dir", str(tmp_path))
    titulos = []
    original = app.m.plt.savefig
    monkeypatch.setattr(app.m.plt, "savefig", lambda *a, **k: (titulos.append(app.m.plt.gcf().axes[0].get_title()), original(*a, **k)))

    app.run(app.m.set_plot_labels, "5", "Titulo X", "", "")  # [5] = Área
    app.run(app.m.plot_area, "Salario")

    assert titulos == ["Titulo X"] and len(list(tmp_path.glob("area_*.png"))) == 1


def test_grafico_com_barra_de_cores_muda_o_eixo_principal(app, monkeypatch):
    app.dummy()
    vistos = _capturar(app, monkeypatch)

    app.run(app.m.set_plot_labels, "8", "Correlações", "", "")  # [8] = Heatmap
    app.run(app.m.plot_correlation_heatmap, "Idade,Salario")

    assert vistos[0][0][0] == "Correlações"


def test_grafico_circular_so_recebe_o_titulo(app, monkeypatch):
    app.dummy()
    vistos = _capturar(app, monkeypatch)

    app.run(app.m.set_plot_labels, "3", "Por cidade", "Não usado", "Nem este")  # [3] = Circular
    app.run(app.m.plot_pie, "Cidade", "Salario", "1")

    assert vistos[0][0] == ("Por cidade", "", "")


def test_texto_do_grafico_e_limpo(app):
    limpo = app.m.clean_plot_text("  Preço  em $ e $x$  ")
    assert limpo == r"Preço em \$ e \$x\$"
    assert len(app.m.clean_plot_text("x" * 500)) == app.m.PLOT_TEXT_MAX


def test_texto_com_cifrao_nao_rebenta_o_grafico(app, monkeypatch):
    app.dummy()
    monkeypatch.setattr(app.m.plt, "show", lambda: app.m.plt.gcf().canvas.draw())

    app.run(app.m.set_plot_labels, "0", "Custo $ total $x^$", "", "")
    app.run(app.m.plot_area, "Salario")


def test_menu_dos_graficos_mostra_quantos_textos_ha_guardados(app, capsys):
    app.dummy()
    app.run(app.m.set_plot_labels, "0", "T", "", "")

    app.run(app.m.menu_charts, "0")

    assert "(1 guardado(s) nesta folha)" in capsys.readouterr().out
