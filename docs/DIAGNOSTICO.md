# Diagnóstico — TermiFolha v1.0

> **Estado actual: todos os problemas abaixo estão resolvidos** (ver a tabela «Estado após as correcções»); as únicas excepções são decisões do dono do projecto (M5 e B3). Este documento é o registo da auditoria à v1.0 e não é actualizado com o código.

> Fase 2. Resultados reais obtidos em Python 3.13.12 / pandas 3.0.2 / numpy 2.4.4 / matplotlib 3.10.9 / python-docx 1.2.0 / fpdf 1.7.2 (Windows 11).
> **[R]** = reproduzido a correr o código. **[L]** = encontrado por leitura, não reproduzido.

## Resultados das ferramentas

| Verificação | Resultado |
|---|---|
| `python -m py_compile` | OK |
| `pip-audit` | sem vulnerabilidades nas dependências do projecto (só avisou de pacotes conda que não existem no PyPI) |
| `ruff` (regras alargadas da configuração do ambiente) | 33 avisos: 13 `except Exception` genérico, 13 `global` desnecessário, 4 `except: pass`, 3 menores |
| Testes | **não existem** |
| 10 gráficos com a coluna `Idade` (inteiro com valor em falta) | todos correm sem erro [R] |
| Exportar Excel, Word e PDF (folha dummy) | funcionam [R] |

## Problemas, por severidade

### ALTO

**A1. Fillna com média ou mediana numa coluna inteira rebenta o programa** [R]
`dados_em_falta`, opções [1]→[2] e [1]→[3]. Na folha dummy, `Idade` é `Int64` e a média é 31.57. Resultado: `TypeError: Invalid value '31.571428571428573' for dtype 'Int64'`, sem `try/except`, o que fecha a aplicação. Impacto: o fluxo mais óbvio de limpeza de dados falha no exemplo oficial. Esforço: baixo.

**A2. Colunas de data criadas com "Nova folha" ficam como `object`** [R]
Depois de `nova_folha` + `adicionar_linha`, a coluna do tipo data tem dtype `object`, porque `aplicar_tipos` só trata int e float. Consequências: o filtro de data nunca é usado (cai no filtro de texto), as estatísticas tratam a data como texto e o groupby e os gráficos trabalham sobre objectos. Só a folha dummy e a importação criam datas correctas. Esforço: baixo.

**A3. Cabeçalhos numéricos em ficheiros Excel provocam `TypeError`** [R]
`pd.read_excel` devolve colunas inteiras (por exemplo anos `2020`, `2021`). `pedir_coluna`, `remover_coluna` e `renomear_coluna` fazem `', '.join(tabela.columns)` e rebentam. Afecta também `editar_celula` e `pedir_coluna_numerica`. Esforço: baixo (normalizar os nomes para `str` na importação).

### MÉDIO

**M1. Erros de gravação e de leitura do estado são engolidos** [R/L]
`guardar_estado` e `carregar_estado` fazem `except Exception: pass`. Com um `.pkl` corrompido o utilizador vê uma folha vazia sem aviso, e o ficheiro é depois sobrescrito. Uma falha de gravação (disco cheio, sem permissões) passa despercebida e os dados perdem-se silenciosamente.

**M2. `pickle` para persistência** [L]
Carregar um `.termifolha_estado.pkl` adulterado executa código arbitrário. Numa aplicação local o risco é baixo, mas o ficheiro vive na pasta de trabalho, que pode ser partilhada.

**M3. Ficheiro de estado relativo ao directório actual** [L]
`FICHEIRO_ESTADO` é um caminho relativo. O manual diz que fica "na mesma pasta" do programa. Quem executa `python D:\...\TermiFolha.py` a partir de outra pasta cria e carrega estados diferentes.

**M4. Adicionar coluna não aceita valor em branco** [R]
O manual diz que se pode deixar em branco (fica `--`). No código, uma entrada vazia é rejeitada em inteiro, decimal e data, e fica num ciclo até se escrever `sair`, que cancela tudo.

**M5. Limpar folha contradiz o manual** [R]
O manual diz que é irreversível e que apaga o histórico de desfazer. Na prática o passo fica no histórico e o Desfazer repõe a folha e recria o `.pkl`.

**M6. Filtro de texto com operador inválido filtra por "contém" sem avisar** [L]
`filtro_texto` trata qualquer opção desconhecida como "contém". Os filtros numérico e de data validam e avisam.

**M7. Sem testes, sem dependências declaradas, sem `.gitignore`, sem README** [R]
Qualquer alteração futura é cega. `fpdf` e `fpdf2` expõem o mesmo módulo `fpdf`, e o ambiente analisado tem o antigo (1.7.2) em vez do `fpdf2` exigido.

### BAIXO

**B1. Divergências manual ↔ código** (dummy com 8 linhas em vez de 10, adicionar linhas em ciclo, remover linha sem mostrar o conteúdo, amostra sem percentagem, describe e únicos de uma coluna só). Pormenores em `docs/MAPA-PROJECTO.md`, secção 8.

**B2. Importar Excel substitui a folha sem confirmar** [L]. Recuperável com Desfazer (até 4 passos).

**B3. Monólito de 2100 linhas com estado global** [L]. Mistura `input`, `print`, cálculo e persistência, o que obriga a simular o teclado para testar.

**B4. Exportação PDF falha com caracteres fora de Latin-1 e usa API deprecada** [R, reclassificado de [L] depois de testar com fpdf2 2.9.0 num venv limpo]. Um valor como `Zażółć` aborta toda a exportação (com `fpdf2` **e** com a `fpdf` 1.7.2; o erro é apanhado e mostrado como "Erro: …"). Com `fpdf2`, a fonte `arial` e `cell(ln=True)` emitem `DeprecationWarning`. Severidade real: MÉDIO.

> **Correcção a este documento:** numa versão anterior escrevi que a exportação com Unicode "funciona na biblioteca instalada". Estava errado: a sonda só verificava se havia excepção e o `exportar_pdf` apanha-a e imprime "Erro". O teste de regressão `test_exportar_pdf_com_caracteres_fora_de_latin1_nao_falha` expôs o problema.

**B5. Avisos do ruff** (`global` onde só se lê, `except` genéricos).

## Pontos verificados que NÃO são problemas
`editar_celula` preserva o dtype `Int64`. Os gráficos lidam bem com valores em falta. A exportação PDF com acentos do português (`ã`, `ç`, `é`, que existem em Latin-1) funciona. A importação de datas e texto com valores nulos infere os tipos correctamente.

## Descoberto durante as correcções (não estava na lista inicial)

- **`filtro_texto` interpretava o termo como expressão regular** [R]: pesquisar `(` dava `PatternError`. Corrigido com `regex=False` (junto com M6).
- **`inplace=True`** [R com pandas 3.1]: `Pandas4Warning`, será removido. 5 ocorrências, corrigidas.
- **Importação (Excel) deixava o estado global meio alterado** se a inferência de tipos falhasse a meio [L]. Corrigido em `adoptar_dataframe`.
- **Inteiros com células vazias importados do Excel ficam "decimal"** [R]: o pandas converte-os para `float64` e não há forma fiável de os distinguir de decimais inteiros (`2000.0`). **Corrigido** perguntando ao utilizador (só quando há ambiguidade); o CSV usa inferência própria.

## Estado após as correcções

| Item | Estado | Commit / nota |
|---|---|---|
| A1 fillna média em inteiro | corrigido | `fix(A1)` |
| A2 datas como `object` | corrigido | `fix(A2)` |
| A3 cabeçalhos numéricos | corrigido | `fix(A3)` |
| M1 erros engolidos | corrigido | `fix(M1,M3)` |
| M2 `pickle` | corrigido: JSON atómico + migração com leitura restrita | `feat(M2)` |
| M3 ficheiro de estado relativo | corrigido | `fix(M1,M3)` |
| M4 coluna em branco | corrigido | `fix(M4)` |
| M5 limpar folha vs manual | **decisão por omissão**: mantém-se reversível com Desfazer; o manual é que está desactualizado | ver README |
| M6 filtro de texto | corrigido | `fix(M6)` |
| M7 testes, dependências, `.gitignore`, README | corrigido | 53 testes, `requirements*.txt`, `pyproject.toml` |
| B1 manual ↔ código | corrigido: remover linha, amostra por percentagem, describe/únicos de todas as colunas; manual reescrito em `MANUAL.md` | `fix(B1)`, `feat` |
| B2 importar sem confirmar | corrigido | `feat` |
| B3 monólito | **por decisão do dono, a aplicação mantém-se num único ficheiro**; identificadores traduzidos | `refactor` |
| B4 PDF | corrigido | `fix(B4)` |
| B5 avisos do ruff | corrigido: 11 `global` desnecessários removidos e 5 `except Exception` estreitados para `ValueError`/`TypeError`/`OverflowError`; os 2 que têm de ser genéricos (ler o estado, avaliar fórmulas) estão justificados com `noqa` | lint limpo com a configuração do projecto e com `BLE`/`PLW0602` |

## Ordem de ataque
A1 → A2 → A3 → M1/M3 → M4 → M5/M6 → M7 (testes, dependências, `.gitignore`, README) → B3 (refactor com testes já a proteger).
