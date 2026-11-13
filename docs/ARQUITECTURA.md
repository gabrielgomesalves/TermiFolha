# Arquitectura

Aplicação de ficheiro único (`TermiFolha.py`), procedural, com estado global. Descrição completa dos módulos e fluxos em `MAPA-PROJECTO.md`; esta página regista as decisões e os pontos de extensão.

## Estado

| Variável | Conteúdo |
|---|---|
| `tabela` | `pandas.DataFrame` activo |
| `tipos` | `dict` coluna → tipo (`str`, `int`, `float`, `"date"`). O tipo vive aqui, não no DataFrame |
| `undo`, `redo` | listas de `(DataFrame, dict)`; limite `HIST_LIMITE = 4` |
| `FICHEIRO_ESTADO` | `.termifolha_estado.pkl`, na pasta do programa |

## Invariantes (o que `aplicar_tipos` garante)

- Colunas `int` → `Int64` (aceita valores em falta); `float` → `float64`; `data` → `datetime64`; `texto` → como vier.
- Toda a operação que altera dados termina com `guardar_estado()` e `registar_ponto()`.
- Nomes de colunas são sempre `str` (a importação converte cabeçalhos numéricos).

## Fluxo de importação

```
importar() / importar_csv()
   └─ lê ficheiro → DataFrame
       └─ adoptar_dataframe(df, detectar_datas)
            ├─ nomes de colunas → str
            ├─ (CSV) texto DD/MM/AAAA → datetime
            ├─ tipos = inferir_tipo por coluna
            └─ só aqui altera tabela/tipos → aplicar_tipos → guardar → registar_ponto
```

O CSV não usa a inferência do pandas: lê tudo como texto (`dtype=str`) e `_numerica_ou_texto` decide, para manter inteiros com falhas como `Int64` e códigos como `007` como texto.

## Exportação PDF

Só fontes core (Latin-1). `_latin1` substitui o resto por `?` e `_pdf_linha` evita `cell(ln=True)`, de modo a funcionar em `fpdf` 1.7 e `fpdf2` sem avisos.

## Testes

`tests/conftest.py` fornece a fixture `app`: isola o estado global, aponta o ficheiro de estado para uma pasta temporária e substitui `input()`. Um teste escreve as respostas do utilizador em sequência (`app.run(app.m.funcao, "resp1", "resp2")`). Se a aplicação pedir mais entradas do que as fornecidas, o teste falha com a pergunta em causa.

## Dívida técnica conhecida

Ver `ROADMAP.md`, itens 1 (`pickle`) e 2 (separar lógica de interface).
