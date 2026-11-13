# Mapa do projecto — TermiFolha v1.0

> Documento da Fase 1 (descoberta). Estado verificado a partir do código e do manual em PDF, **antes de qualquer alteração (v1.0)**. É um retrato histórico: os nomes de funções e variáveis são os originais em português (correspondência em `RENOMEACOES.md`) e a persistência por `pickle` já foi substituída por JSON. Para o estado actual, ver `ARQUITECTURA.md`.

## 1. Resumo

TermiFolha é uma folha de cálculo em linha de comando, em Python. Permite criar folhas com colunas tipadas, editar linhas e colunas, ordenar, filtrar, calcular estatísticas, desenhar gráficos e importar ou exportar para Excel, Word e PDF. O estado é gravado automaticamente entre sessões. O autor é Gabriel Gomes Alves, a licença é MIT.

## 2. Árvore de ficheiros (estado inicial)

```
TermiFolha/
├── TermiFolha.py                                  # toda a aplicação (2100 linhas, 83 funções)
└── Manual do Utilizador — TermiFolha v1.0.pdf     # manual (14 secções + anexo)
```

Em tempo de execução cria também `.termifolha_estado.pkl` (estado da sessão) na pasta de trabalho actual.

Não existem: README, testes, `requirements.txt`/`pyproject.toml`, `.gitignore`, CI, Docker, `.env.example`. O projecto não usa variáveis de ambiente, rede nem segredos.

## 3. Stack

| Componente | Notas |
|---|---|
| Python | 3.13.12 (manual indica "testado em 3.13") |
| pandas 3.0.2, numpy 2.4.4 | tabela e cálculo |
| matplotlib 3.10.9 | gráficos em janela própria |
| openpyxl 3.1.5 | leitura e escrita de `.xlsx` (usado indirectamente pelo pandas) |
| python-docx 1.2.0 | exportação Word |
| **fpdf 1.7.2** | instalado no ambiente o pacote antigo `fpdf`, não o `fpdf2` que o manual e o código dizem exigir. Ambos expõem `from fpdf import FPDF` |
| pickle (stdlib) | persistência do estado |

## 4. Arquitectura

Um único módulo procedural com estado global:

```
main()
 ├─ carregar_estado()  ← .termifolha_estado.pkl
 └─ ciclo de menu
     ├─ menu_folha     → nova_folha / importar / exportar / exportar_word / exportar_pdf / criar_folha_dummy / limpar_folha
     ├─ menu_linhas    → adicionar_linha / editar_celula / remover_linha
     ├─ menu_colunas   → adicionar_coluna / remover_coluna / renomear_coluna
     ├─ menu_analise   → ordenar / filtrar (numérico, data, texto) / ver_inicio_fim / amostra
     ├─ estatisticas   → 8 funções estat_*
     ├─ menu_avancado  → info_resumo / agrupar / duplicados / dados_em_falta
     ├─ menu_graficos  → 10 funções gerar_*
     └─ sobre / desfazer / refazer
```

**Estado global** (`TermiFolha.py:78-82`): `tabela` (DataFrame), `tipos` (dict coluna→tipo), `undo` e `redo` (listas de tuplos `(DataFrame, dict)`).

**Fluxo de dados típico de uma operação que altera dados:**
`input()` → validação/`converter()` → alteração global em `tabela` → `aplicar_tipos()` (nem sempre) → `guardar_estado()` (pickle) → `registar_ponto()` (undo, máximo 4).

## 5. Mapa de módulos (por secção do ficheiro)

| Linhas | Responsabilidade |
|---|---|
| 53–82 | Constantes, tipos suportados, estado global |
| 85–117 | Utilitários de terminal e formatação (`limpar`, `pausar`, `truncar`, `formatar_valor`, `formatar_numero`) |
| 119–191 | Desenho da tabela e cabeçalho |
| 193–346 | Entrada do utilizador: opções, tipos, conversão, resolução de colunas, prompts |
| 348–415 | `aplicar_tipos`, persistência, undo/redo |
| 417–702 | Gestão da folha: criar, importar, exportar (Excel/Word/PDF), dummy, limpar |
| 705–850 | Operações sobre linhas |
| 852–973 | Operações sobre colunas |
| 976–1172 | Análise básica |
| 1175–1363 | Estatísticas |
| 1366–1535 | Análise avançada |
| 1538–1858 | Gráficos |
| 1861–2043 | Menus e ecrã "Sobre" |
| 2046–2099 | `main()` |

## 6. Inventário de funcionalidades existentes

- **Folha:** criar com colunas tipadas (texto, inteiro, decimal, data), importar `.xlsx` com inferência de tipos, exportar `.xlsx`, `.docx` e `.pdf`, folha dummy de 8 linhas, limpar.
- **Linhas:** adicionar em ciclo (termina com `sair`), editar célula, remover com confirmação.
- **Colunas:** adicionar (com preenchimento linha a linha), remover, renomear.
- **Análise básica:** ordenar por várias colunas, filtrar por tipo, primeiras ou últimas N linhas, amostra aleatória.
- **Estatísticas:** describe, base, desvio e variância, moda, quantis, top N, únicos, assimetria e curtose.
- **Avançada:** info, groupby com 5 agregações, duplicados (detectar e remover), dados em falta (fillna ou dropna).
- **Gráficos (10):** barras, linhas, circular, dispersão, área acumulada, histograma, boxplot, heatmap de correlação, média móvel, tendência polinomial.
- **Sessão:** undo/redo com limite de 4 passos, gravação automática.

## 7. Como correr

```
pip install pandas numpy matplotlib openpyxl python-docx fpdf2
python TermiFolha.py
```

Não há testes, lint nem build definidos.

## 8. Incoerências detectadas entre manual e código

| # | Manual | Código |
|---|---|---|
| 1 | Folha dummy com 10 linhas | cria 8 linhas (`TermiFolha.py:655-664`) |
| 2 | Adicionar linhas "pede quantas linhas adicionar" | ciclo contínuo até `sair` |
| 3 | Remover linha mostra o conteúdo antes de confirmar | só pede confirmação |
| 4 | Adicionar coluna: preenchimento pode ficar em branco (`--`) | entrada vazia é rejeitada em inteiro, decimal e data |
| 5 | Amostra aleatória por percentagem ou número fixo | só número |
| 6 | Describe resume todas as colunas numéricas | pede uma coluna |
| 7 | Únicos "por coluna" | pede uma coluna |
| 8 | Limpar folha apaga o histórico de undo | `limpar_folha` chama `registar_ponto()`, e o Desfazer repõe a folha |
| 9 | Dependência `fpdf2` | o ambiente tem `fpdf` 1.7.2 |

## 9. Observações estruturais

- Monólito de 2100 linhas, sem separação entre lógica de dados e interface (cada função mistura `input`, `print`, cálculo e persistência). Isto impede testes unitários sem simular o teclado.
- Estado global mutável acedido por todas as funções.
- Identificadores em português (`tabela`, `adicionar_linha`, …). O pedido pede identificadores em inglês, mas renomear tudo quebra a convenção actual do projecto e gera um diff enorme. Decisão por omissão: ver `docs/DIAGNOSTICO.md` e o relatório final.
- Persistência por `pickle`, que executa código ao carregar um ficheiro adulterado.

## 10. Glossário

- **Folha / tabela:** o DataFrame activo.
- **Tipo:** um de texto, inteiro, decimal, data. Guardado em `tipos`, não no DataFrame.
- **Ponto de undo:** cópia de `(tabela, tipos)`, guardada em `undo` por `registar_ponto()`.
- **Estado:** ficheiro `.termifolha_estado.pkl` com `(tabela, tipos)`.
