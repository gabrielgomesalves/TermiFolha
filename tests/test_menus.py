"""Cada opção de cada menu tem de chamar a função certa (apanha opções anunciadas mas não ligadas)."""
import re

import pytest

# menu -> {opção: função chamada}
_MENUS = {
    "menu_sheet": {
        "1": "new_sheet", "2": "import_excel", "3": "export_excel", "4": "export_word", "5": "export_pdf", "6": "create_dummy_sheet",
        "7": "clear_sheet", "8": "import_csv", "9": "export_csv", "A": "import_json", "B": "import_pasted", "C": "export_json",
        "D": "export_markdown", "E": "export_html", "F": "import_url", "G": "import_database", "P": "save_project", "O": "open_project",
    },
    "menu_rows": {"1": "add_row", "2": "edit_cell", "3": "remove_row", "4": "find_replace"},
    "menu_columns": {
        "1": "add_column", "2": "remove_column", "3": "rename_column", "4": "format_column", "5": "choose_visible_columns",
        "6": "add_formula_column", "7": "freeze_formula", "8": "convert_column_type", "9": "menu_rules", "A": "normalize_column", "B": "bin_column",
    },
    "menu_basic_analysis": {"1": "sort_rows", "2": "filter_rows", "3": "view_head_tail", "4": "random_sample", "5": "view_pages", "6": "choose_totals"},
    "menu_advanced": {"1": "info_summary", "2": "group_by", "3": "duplicates", "4": "missing_data", "5": "pivot_table_ui", "6": "outliers"},
    "menu_charts": {
        "1": "plot_bars", "2": "plot_lines", "3": "plot_pie", "4": "plot_scatter", "5": "plot_area", "6": "plot_histogram", "7": "plot_boxplot",
        "8": "plot_correlation_heatmap", "9": "plot_moving_average", "A": "plot_polynomial_trend", "B": "toggle_png_export", "C": "set_plot_labels",
    },
    "menu_workbook": {
        "M": "switch_sheet_ui", "A": "add_blank_sheet", "R": "rename_sheet_ui", "E": "delete_sheet_ui", "D": "duplicate_sheet_ui",
        "J": "merge_sheets_ui", "C": "concat_sheets_ui",
    },
    "menu_rules": {"1": "set_rule", "2": "remove_rule", "3": "check_rules"},
    "statistics_menu": {
        "1": "stat_summary", "2": "stat_basic", "3": "stat_std_var", "4": "stat_mode", "5": "stat_quantiles", "6": "stat_counts",
        "7": "stat_unique", "8": "stat_skew_kurt", "9": "stat_correlation",
    },
}
_CASOS = [(menu, opcao, alvo) for menu, opcoes in _MENUS.items() for opcao, alvo in opcoes.items()]


@pytest.mark.parametrize("menu, opcao, alvo", _CASOS, ids=[f"{m}:{o}" for m, o, _ in _CASOS])
def test_opcao_do_menu_chama_a_funcao_certa(app, monkeypatch, menu, opcao, alvo):
    app.dummy()  # alguns menus (estatísticas) só abrem com dados
    chamadas = []
    monkeypatch.setattr(app.m, alvo, lambda: chamadas.append(alvo))

    app.run(getattr(app.m, menu), opcao, "0")

    assert chamadas == [alvo]


@pytest.mark.parametrize("opcao, alvo", [
    ("1", "menu_sheet"), ("2", "menu_rows"), ("3", "menu_columns"), ("4", "menu_basic_analysis"), ("5", "statistics_menu"),
    ("6", "menu_advanced"), ("7", "menu_charts"), ("8", "menu_workbook"), ("9", "menu_settings"), ("S", "about"), ("U", "undo_last"), ("R", "redo_last"),
])
def test_opcao_do_menu_principal_chama_a_funcao_certa(app, monkeypatch, opcao, alvo):
    chamadas = []
    monkeypatch.setattr(app.m, alvo, lambda: chamadas.append(alvo))

    app.run(app.m.run_interactive, opcao, "Q")

    assert chamadas == [alvo]


def test_opcoes_anunciadas_e_ligadas_coincidem(app, capsys):
    """O texto de cada menu lista as opções [x]; têm de ser exactamente as que o mapa acima liga."""
    app.dummy()
    for menu, opcoes in _MENUS.items():
        capsys.readouterr()
        app.run(getattr(app.m, menu), "0")
        anunciadas = set(re.findall(r"\[([0-9A-Z])\]", capsys.readouterr().out)) - {"0"}
        if menu == "menu_workbook":
            anunciadas = {a for a in anunciadas if a.isalpha()}  # [1], [2]... são os números das folhas
        assert anunciadas == set(opcoes), menu
