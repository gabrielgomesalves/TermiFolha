# Arquitectura

Aplicação de ficheiro único (`TermiFolha.py`), procedural, com estado global. Por decisão do dono do projecto mantém-se assim. Os nomes em inglês; ver `RENOMEACOES.md` para a correspondência com a v1.0. `MAPA-PROJECTO.md` é o retrato inicial (nomes antigos).

## Estado

| Variável | Conteúdo |
|---|---|
| `table` | `pandas.DataFrame` da folha **activa** |
| `col_types` | `dict` coluna → tipo (`str`, `int`, `float`, `DATE_TYPE`). O tipo vive aqui, não no DataFrame |
| `meta` | metadados da folha activa (`new_meta()`): `rules` (validação), `formats`, `visible`, `formulas`, `totals` |
| `sheets` | `dict` nome → folha, na ordem dos separadores. As folhas inactivas são `{"table", "types", "meta"}`; a **activa é `None`** (vive em `table`/`col_types`/`meta`). `active_name` diz qual é |
| `settings` | `history_limit` (omissão 20, máx. 100) e `page_size` |
| `undo_stack`, `redo_stack` | listas de `snapshot()` (o livro inteiro: folha activa, outras folhas, nome da activa) |
| `STATE_ENABLED` | `False` no modo de linha de comandos: o estado da sessão não é lido nem gravado |
| `plot_png_dir` | pasta de gravação dos gráficos em PNG; `None` = abrir janela |
| `STATE_FILE` | `.termifolha_estado.json`, na pasta do programa |
| `LEGACY_STATE_FILE` | `.termifolha_estado.pkl` (v1.0), só lido para migrar |
| `LOG_FILE` | `.termifolha.log`, só criado por `main()` |

## Invariantes (o que `apply_types` garante)

- `int` → `Int64` (aceita valores em falta); `float` → `float64`; data → `datetime64`; texto → como vier.
- Toda a operação que altera dados termina com `commit()` = `recalculate_formulas()` + `save_state()` + `record_checkpoint()`.
- Nomes de colunas são sempre `str`. Nomes de folhas: únicos (sem distinguir maiúsculas), até 31 caracteres, sem `[]:*?/\` (`sheet_name_error`).
- `meta` acompanha as colunas: `rename_column_meta`/`forget_column_meta`. Uma coluna usada por uma fórmula não se renomeia, remove nem converte (`formula_users`).
- Mudar de folha (`switch_sheet`) só troca as referências globais; é um passo do histórico.

## Persistência

```
save_state()  → _state_to_dict() → JSON → ficheiro .tmp → os.replace()      (atómico)
load_state()  → JSON válido?  → _state_from_dict() → apply_types()
              → senão, só há .pkl? → _SafeUnpickler → migra, grava JSON, renomeia para .pkl.migrado
              → ficheiro inválido → preservado como .corrompido, sessão vazia, utilizador avisado
```

Formato JSON **versão 2**: `{"versao": 2, "activa": nome, "definicoes": {...}, "folhas": [{"nome", "colunas", "tipos": {col: "str|int|float|date"}, "dados": [[...]], "regras", "formatos", "visiveis", "formulas", "totais"}]}`; datas em ISO 8601; valores em falta `null`. As chaves ficam em português para que os ficheiros já gravados continuem válidos. A **versão 1** (uma folha, sem metadados) continua a ser lida e convertida. `_state_from_dict` valida a estrutura (versão, tipos, comprimento de cada linha, forma dos metadados, nomes de folhas) antes de aceitar o conteúdo; qualquer falha preserva o ficheiro como `.corrompido`.

Um **projecto** (`.termifolha`) é o mesmo JSON (`_write_state_file`/`_read_state_file`), mas as definições do ficheiro não são importadas ao abrir.

**Leitura restrita do `.pkl`:** `_SafeUnpickler.find_class` só devolve classes de `pandas`/`numpy` necessárias para reconstruir um DataFrame, `builtins` inócuos e nada mais; recusa nomes com ponto e `pandas.core.computation`. Testado com payloads que tentam `os.system` e `pandas.eval`.

## Fórmulas

`evaluate_formula` nunca usa `eval`/`exec`. `_parse_formula` faz `ast.parse(..., mode="eval")` (depois de trocar `[Nome com espaços]` por identificadores seguros) e **só aceita** os nós de uma lista fechada (`_FORMULA_NODES`), constantes numéricas/texto, nomes de colunas e as funções de `_FORMULA_ARITY` (sem atributos, subscritos, lambdas, compreensões nem argumentos nomeados); limita o tamanho (500 caracteres, 200 nós). Os números literais são avaliados como `float64` (evita potências gigantes em inteiros Python) e os infinitos viram em falta. Os operandos são séries: números como `float64`, datas como `datetime64`, texto como `object`. O tipo do resultado é inferido (`int` só se todas as colunas usadas forem inteiras/datas, não houver `/`, `sqrt`, `log`, `exp` e o resultado for integral). `recalculate_formulas` percorre `meta["formulas"]` pela ordem de criação (uma fórmula só pode usar colunas anteriores, por isso não há ciclos); uma fórmula que falhe deixa a coluna em branco e avisa.

## Livro de folhas

`add_sheet`, `switch_sheet`, `rename_sheet`, `delete_sheet`, `duplicate_sheet` são o núcleo (testável, sem `input()`); as funções `*_ui` são só diálogo. `merge_tables` (renomeia colunas repetidas, verifica os tipos das chaves, estima as linhas antes de juntar), `concat_tables` e `pivot_data` devolvem `(DataFrame, tipos)`; a interface cria sempre uma folha nova. `adopt_workbook` instala várias folhas de uma vez (Excel com vários separadores), sem deixar o livro a meio se algo falhar.

## Entrada e saída

Cada formato tem uma função `write_*`/`read_*` sem interacção (`write_excel/csv/json/markdown/html/word/pdf`, `read_json_file`, `parse_csv_text`, `_read_csv`); as funções `export_*`/`import_*` só pedem o nome e tratam erros. O modo de linha de comandos (`main(argv)` → `run_batch`) usa as primeiras e devolve códigos de saída; nunca chama `input()`. `write_html` escapa todo o texto com `html.escape` e não emite scripts.

## Fluxo de importação

```
import_excel() / import_csv()
   └─ confirm_replace_sheet()  (se já existir folha)
       └─ lê ficheiro → DataFrame
           └─ adopt_dataframe(df, detect_dates, ask_integers)
                ├─ nomes de colunas → str
                ├─ (Excel) colunas só com inteiros e células vazias → pergunta ao utilizador
                ├─ (CSV) texto DD/MM/AAAA → datetime
                ├─ col_types = infer_type por coluna
                └─ só aqui altera table/col_types → apply_types → save_state → record_checkpoint
```

O CSV não usa a inferência do pandas: lê tudo como texto (`dtype=str`) e `_numeric_or_text` decide (inteiros com falhas → `Int64`, `1,5` → decimal, `007` → texto).

## Exportação PDF

`_configure_pdf_font` regista a DejaVu Sans (incluída no matplotlib) quando o `fpdf` é o `fpdf2`; com o `fpdf` 1.7 mantém Helvetica porque este escreve ficheiros de cache na pasta da fonte. `_pdf_text` reduz a Latin-1 quando não há fonte Unicode e `_pdf_line` evita o `cell(ln=True)` deprecado.

## Registo

`log = logging.getLogger("termifolha")` com `NullHandler`: importar o módulo não cria ficheiros. `configure_log()` (chamada por `main()`) acrescenta um `RotatingFileHandler` de 256 KB. Os blocos `except` de importação, exportação e gravação usam `log.exception`.

## Testes

`tests/conftest.py` fornece a fixture `app`: isola o estado global, aponta os ficheiros de estado para uma pasta temporária e substitui `input()`. Um teste escreve as respostas do utilizador em sequência (`app.run(app.m.add_column, "Nome", "1", ...)`). Se a aplicação pedir mais entradas do que as fornecidas, o teste falha com a pergunta em causa.
