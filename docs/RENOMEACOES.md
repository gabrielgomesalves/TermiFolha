# Renomeações (português → inglês)

Os identificadores do código passaram para inglês na versão 1.1. Esta tabela serve para quem tenha scripts ou notas com os nomes antigos. Texto apresentado ao utilizador, comentários e documentação continuam em português.

**Mantido em português de propósito:** as chaves do ficheiro de estado (`versao`, `colunas`, `tipos`, `dados`), para que os ficheiros já gravados continuem a abrir.

## Estado global

| Antes | Agora |
|---|---|
| `redo` | `redo_stack` |
| `tabela` | `table` |
| `tipos` | `col_types` |
| `undo` | `undo_stack` |

## Constantes

| Antes | Agora |
|---|---|
| `AUTOR_APP` | `APP_AUTHOR` |
| `COL_LIMITE` | `COLUMN_WIDTH_LIMIT` |
| `ETIQUETAS_ESTAT` | `STAT_LABELS` |
| `FICHEIRO_ESTADO` | `STATE_FILE` |
| `FICHEIRO_ESTADO_LEGADO` | `LEGACY_STATE_FILE` |
| `FICHEIRO_LOG` | `LOG_FILE` |
| `HIST_LIMITE` | `HISTORY_LIMIT` |
| `LARGURA` | `WIDTH` |
| `NOME_APP` | `APP_NAME` |
| `NOME_TIPO` | `TYPE_NAME` |
| `TIPOS` | `TYPE_CHOICES` |
| `TIPO_DATA` | `DATE_TYPE` |
| `TITULO_APP` | `APP_TITLE` |
| `UI_INDENTACAO_1` | `UI_INDENT_1` |
| `UI_INDENTACAO_2` | `UI_INDENT_2` |
| `VERSAO_APP` | `APP_VERSION` |
| `_NOME_TIPO_JSON` | `_JSON_TYPE_NAME` |
| `_PASTA_PROGRAMA` | `_PROGRAM_DIR` |
| `_TIPO_DE_NOME_JSON` | `_TYPE_FROM_JSON_NAME` |
| `_TODAS` | `_ALL` |
| `_VERSAO_ESTADO` | `_STATE_VERSION` |

## Funções

| Antes | Agora |
|---|---|
| `_como_datas` | `_as_dates` |
| `_configurar_fonte_pdf` | `_configure_pdf_font` |
| `_estado_de_dict` | `_state_from_dict` |
| `_estado_para_dict` | `_state_to_dict` |
| `_inteiros_com_falhas` | `_integers_with_gaps` |
| `_ler_csv` | `_read_csv` |
| `_numerica_ou_texto` | `_numeric_or_text` |
| `_pdf_linha` | `_pdf_line` |
| `_preservar_ficheiro_invalido` | `_preserve_invalid_file` |
| `_resumo_todas_numericas` | `_summary_all_numeric` |
| `_separador_csv` | `_csv_separator` |
| `_texto_pdf` | `_pdf_text` |
| `_valor_json` | `_json_value` |
| `adicionar_coluna` | `add_column` |
| `adicionar_linha` | `add_row` |
| `adoptar_dataframe` | `adopt_dataframe` |
| `agrupar` | `group_by` |
| `amostra` | `random_sample` |
| `aplicar_tipos` | `apply_types` |
| `cabecalho` | `header` |
| `carregar_estado` | `load_state` |
| `configurar_log` | `configure_log` |
| `confirmar` | `confirm` |
| `confirmar_substituir_folha` | `confirm_replace_sheet` |
| `converter` | `convert_value` |
| `criar_folha_dummy` | `create_dummy_sheet` |
| `dados_em_falta` | `missing_data` |
| `desenhar_tabela` | `draw_table` |
| `desenhar_tabela_dados` | `draw_data_table` |
| `desfazer` | `undo_last` |
| `duplicados` | `duplicates` |
| `editar_celula` | `edit_cell` |
| `estat_asim_kurt` | `stat_skew_kurt` |
| `estat_base` | `stat_basic` |
| `estat_contagem` | `stat_counts` |
| `estat_desv_var` | `stat_std_var` |
| `estat_moda` | `stat_mode` |
| `estat_quantis` | `stat_quantiles` |
| `estat_resumo` | `stat_summary` |
| `estat_unicos` | `stat_unique` |
| `estatisticas` | `statistics_menu` |
| `exportar` | `export_excel` |
| `exportar_csv` | `export_csv` |
| `exportar_pdf` | `export_pdf` |
| `exportar_word` | `export_word` |
| `filtrar` | `filter_rows` |
| `filtro_data` | `filter_date` |
| `filtro_numerico` | `filter_numeric` |
| `filtro_texto` | `filter_text` |
| `formatar_numero` | `format_number` |
| `formatar_valor` | `format_value` |
| `gerar_barras_eixos` | `plot_bars` |
| `gerar_circular_eixos` | `plot_pie` |
| `gerar_diagrama_caixas` | `plot_boxplot` |
| `gerar_grafico_area` | `plot_area` |
| `gerar_grafico_dispersao` | `plot_scatter` |
| `gerar_histograma` | `plot_histogram` |
| `gerar_linhas_eixos` | `plot_lines` |
| `gerar_linhas_media_movel` | `plot_moving_average` |
| `gerar_mapa_calor_correlacao` | `plot_correlation_heatmap` |
| `gerar_tendencia_polinomial` | `plot_polynomial_trend` |
| `guardar_estado` | `save_state` |
| `importar` | `import_excel` |
| `importar_csv` | `import_csv` |
| `imprimir_tabela_pandas` | `print_pandas_table` |
| `inferir_tipo` | `infer_type` |
| `info_resumo` | `info_summary` |
| `larguras` | `column_widths` |
| `limpar` | `clear_screen` |
| `limpar_folha` | `clear_sheet` |
| `menu_analise` | `menu_basic_analysis` |
| `menu_avancado` | `menu_advanced` |
| `menu_colunas` | `menu_columns` |
| `menu_folha` | `menu_sheet` |
| `menu_graficos` | `menu_charts` |
| `menu_linhas` | `menu_rows` |
| `nova_folha` | `new_sheet` |
| `ordenar` | `sort_rows` |
| `pausar` | `pause` |
| `pedir_coluna` | `ask_column` |
| `pedir_coluna_numerica` | `ask_numeric_column` |
| `pedir_coluna_ou_todas` | `ask_column_or_all` |
| `pedir_colunas_numericas` | `ask_numeric_columns` |
| `pedir_inteiro_intervalo` | `ask_int_in_range` |
| `pedir_opcao` | `ask_option` |
| `pedir_tipo` | `ask_type` |
| `refazer` | `redo_last` |
| `registar_ponto` | `record_checkpoint` |
| `remover_coluna` | `remove_column` |
| `remover_linha` | `remove_row` |
| `renomear_coluna` | `rename_column` |
| `resolver_coluna` | `resolve_column` |
| `sobre` | `about` |
| `tipo_de` | `type_of` |
| `truncar` | `truncate` |
| `valor_para_coluna` | `value_for_column` |
| `ver_inicio_fim` | `view_head_tail` |

Os parâmetros e variáveis locais também foram traduzidos; não estão listados porque não são visíveis fora da função.
