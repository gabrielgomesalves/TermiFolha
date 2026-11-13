# Diagnóstico — TermiFolha v1.0

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

**B4. Exportação PDF depende da versão da biblioteca** [L]. Com `fpdf2` recente, `cell(..., ln=True)` está deprecado e caracteres fora de Latin-1 com as fontes core levantam excepção. Com a `fpdf` 1.7.2 testada, nada disto acontece. Não consegui verificar o comportamento com `fpdf2`.

**B5. Avisos do ruff** (`global` onde só se lê, `except` genéricos).

## Pontos verificados que NÃO são problemas
`editar_celula` preserva o dtype `Int64`. Os gráficos lidam bem com valores em falta. A exportação com acentos e Unicode funciona na biblioteca instalada. A importação de datas e texto com valores nulos infere os tipos correctamente.

## Ordem de ataque
A1 → A2 → A3 → M1/M3 → M4 → M5/M6 → M7 (testes, dependências, `.gitignore`, README) → B3 (refactor com testes já a proteger).
