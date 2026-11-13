# Arquitectura

Aplicação de ficheiro único (`TermiFolha.py`), procedural, com estado global. Por decisão do dono do projecto mantém-se assim. Os nomes em inglês; ver `RENOMEACOES.md` para a correspondência com a v1.0. `MAPA-PROJECTO.md` é o retrato inicial (nomes antigos).

## Estado

| Variável | Conteúdo |
|---|---|
| `table` | `pandas.DataFrame` activo |
| `col_types` | `dict` coluna → tipo (`str`, `int`, `float`, `DATE_TYPE`). O tipo vive aqui, não no DataFrame |
| `undo_stack`, `redo_stack` | listas de `(DataFrame, dict)`; limite `HISTORY_LIMIT = 4` |
| `STATE_FILE` | `.termifolha_estado.json`, na pasta do programa |
| `LEGACY_STATE_FILE` | `.termifolha_estado.pkl` (v1.0), só lido para migrar |
| `LOG_FILE` | `.termifolha.log`, só criado por `main()` |

## Invariantes (o que `apply_types` garante)

- `int` → `Int64` (aceita valores em falta); `float` → `float64`; data → `datetime64`; texto → como vier.
- Toda a operação que altera dados termina com `save_state()` e `record_checkpoint()`.
- Nomes de colunas são sempre `str`.

## Persistência

```
save_state()  → _state_to_dict() → JSON → ficheiro .tmp → os.replace()      (atómico)
load_state()  → JSON válido?  → _state_from_dict() → apply_types()
              → senão, só há .pkl? → _SafeUnpickler → migra, grava JSON, renomeia para .pkl.migrado
              → ficheiro inválido → preservado como .corrompido, sessão vazia, utilizador avisado
```

Formato JSON: `{"versao": 1, "colunas": [...], "tipos": {col: "str|int|float|date"}, "dados": [[...]]}`; datas em ISO 8601; valores em falta `null`. As chaves ficam em português para que os ficheiros já gravados continuem válidos. `_state_from_dict` valida a estrutura (versão, tipos de colunas, comprimento de cada linha) antes de aceitar o conteúdo.

**Leitura restrita do `.pkl`:** `_SafeUnpickler.find_class` só devolve classes de `pandas`/`numpy` necessárias para reconstruir um DataFrame, `builtins` inócuos e nada mais; recusa nomes com ponto e `pandas.core.computation`. Testado com payloads que tentam `os.system` e `pandas.eval`.

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
