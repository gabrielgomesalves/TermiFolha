# Manual do Utilizador — TermiFolha v1.1

Este manual substitui o PDF da v1.0, que tinha divergências face ao programa (ver `CHANGELOG.md`). Descreve o comportamento real da versão 1.1.

## 1. Instalação e arranque

```
pip install -r requirements.txt
python TermiFolha.py
```

Requer Python 3.11 ou superior (testado em 3.13). Não instale o pacote antigo `fpdf` ao lado do `fpdf2`: partilham o mesmo nome de módulo. O programa funciona com ambos, mas só o `fpdf2` exporta PDF com caracteres fora de Latin-1.

Ao arrancar, o estado da sessão anterior é restaurado a partir de `.termifolha_estado.json`, que fica **na pasta do programa**. Erros de execução ficam em `.termifolha.log` (na mesma pasta).

## 2. Conceitos

| Conceito | Descrição |
|---|---|
| Folha | A tabela activa. |
| Coluna | Tem um nome único (sem distinguir maiúsculas) e um tipo fixo. |
| Linha | Identificada por um índice que começa em 0. |
| Tipos | texto, inteiro, decimal, data (`DD/MM/AAAA`). |

- Decimais usam **ponto** (`3.14`). A importação de CSV aceita também a vírgula decimal (`3,14`).
- Células em falta aparecem como `--`.
- Em qualquer prompt, escrever `sair` cancela a operação.
- Nos prompts com valor predefinido (`Enter = 5`), Enter aceita-o.

## 3. Menu principal

| Tecla | Acção |
|---|---|
| 1 | Folha: nova, importar, exportar, dummy, limpar |
| 2 | Linhas: adicionar, editar, remover |
| 3 | Colunas: adicionar, remover, renomear |
| 4 | Análise básica: ordenar, filtrar, início/fim, amostra |
| 5 | Estatísticas (8 métodos) |
| 6 | Análise avançada: info, agrupar, duplicados, dados em falta |
| 7 | Gráficos (10 tipos) |
| S | Sobre |
| U / R | Desfazer / Refazer |
| Q | Sair |

## 4. Folha

| Opção | Descrição |
|---|---|
| [1] Nova folha | Pede nome e tipo de cada coluna; Enter num nome termina. |
| [2] Importar Excel | Lê `.xlsx`. Pede confirmação se já existir folha. Se uma coluna só tiver inteiros e células vazias, pergunta se deve ser inteira (o pandas, por omissão, torna-a decimal). |
| [3] Exportar Excel | Grava `.xlsx`. |
| [4] Exportar Word | Colunas e tipos, tabela completa e estatísticas das colunas numéricas. |
| [5] Exportar PDF | O mesmo conteúdo. As células da tabela são truncadas a 14 caracteres. Com `fpdf2` usa a fonte DejaVu Sans (Unicode); com o `fpdf` antigo, caracteres fora de Latin-1 saem como `?`. |
| [6] Folha dummy | 8 linhas e 5 colunas de exemplo, com um valor em falta e uma linha duplicada. Pede confirmação se já existir folha. |
| [7] Limpar folha | Apaga a folha e o ficheiro de estado. **Pode ser desfeita** com `[U]`. |
| [8] Importar CSV | Detecta o separador (`,` `;` tab `|`) e a codificação (UTF-8 ou Latin-1). Reconhece datas `DD/MM/AAAA`, vírgula decimal e mantém códigos como `007` como texto. Pede confirmação se já existir folha. |
| [9] Exportar CSV | Escolha entre formato internacional (`,` e `.`) e Excel português (`;` e `,`). Grava em UTF-8 com BOM. |

A extensão é acrescentada automaticamente se faltar. Pode indicar um caminho completo.

## 5. Linhas e colunas

**Linhas** — Adicionar: introduz linhas em ciclo, valor a valor, até escrever `sair`; os tipos são validados. Editar célula: índice, coluna e novo valor. Remover: mostra a linha e pede confirmação.

**Colunas** — Adicionar: pede nome e tipo e depois um valor por linha existente; **Enter deixa a célula em branco** (`--`). Remover e Renomear pedem o nome da coluna.

## 6. Análise básica

- **Ordenar** por uma ou mais colunas, cada uma ascendente (`A`) ou descendente (`D`).
- **Filtrar** conforme o tipo da coluna: numérico (`== > < >= <=` e intervalo), data (antes, depois, entre) e texto (contém, igual, começa, termina; sem distinguir maiúsculas; o termo é literal, não uma expressão regular). O resultado é só mostrado; a folha não muda.
- **Ver início/fim:** as primeiras ou últimas N linhas.
- **Amostra aleatória:** um número de linhas (`3`) ou uma percentagem (`20%`).

## 7. Estatísticas

| Opção | Resultado |
|---|---|
| [1] Describe | Resumo de uma coluna, ou de **todas as numéricas** com Enter. |
| [2] Base | Média, soma, mínimo, máximo e mediana. |
| [3] Desvio e variância | |
| [4] Moda | |
| [5] Quartis / percentis | Percentis entre 0 e 1. |
| [6] Top N | Valores mais frequentes. |
| [7] Únicos | Uma coluna, ou **todas** com Enter. |
| [8] Assimetria e curtose | |

## 8. Análise avançada

- **Info:** dimensões, memória, tipo, valores não nulos e percentagem em falta por coluna.
- **Agrupar:** agrupa por uma coluna e agrega (soma, média, contagem, mínimo, máximo) uma coluna numérica.
- **Duplicados:** mostra e, se confirmar, remove mantendo a primeira ocorrência.
- **Dados em falta:** preencher (valor, média, mediana, *forward-fill*) ou remover linhas. Em colunas inteiras, média e mediana são **arredondadas**.

## 9. Gráficos

Cada gráfico abre numa janela do matplotlib; feche-a para voltar ao menu.

| | Gráfico | | Gráfico |
|---|---|---|---|
| 1 | Barras (eixos à escolha) | 6 | Histograma com quartis |
| 2 | Linhas (eixo X à escolha) | 7 | Boxplot |
| 3 | Circular | 8 | Heatmap de correlação |
| 4 | Dispersão | 9 | Linhas com média móvel |
| 5 | Área (soma acumulada) | A | Tendência polinomial |

## 10. Desfazer e refazer

Cada alteração guarda um ponto de recuperação; o histórico tem no máximo **4 passos**. `[U]` desfaz e `[R]` refaz. Exporte periodicamente para Excel: o histórico não substitui uma cópia de segurança.

## 11. Resolução de problemas

| Problema | O que fazer |
|---|---|
| Erro ao importar Excel | Confirme a extensão `.xlsx`, que `openpyxl` está instalado e que o ficheiro não está aberto noutra aplicação. |
| Erro ao exportar | Confirme as permissões de escrita e que o ficheiro de destino não está aberto. Detalhes em `.termifolha.log`. |
| O programa avisa que o estado está corrompido | O ficheiro foi preservado como `.termifolha_estado.json.corrompido`; a sessão começa vazia. |
| Estado de uma versão antiga | Um `.termifolha_estado.pkl` é migrado automaticamente no primeiro arranque (cópia em `.pkl.migrado`). |
| O gráfico não aparece | A janela pode ter aberto por trás de outras. |
| Valores cortados no ecrã | Cada coluna é truncada a 15 caracteres no ecrã; os dados estão íntegros. |
| Folha com dados no arranque | É o estado da sessão anterior; use `[1]` → `[7]` para começar do zero. |
