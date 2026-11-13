# Manual do Utilizador — TermiFolha v1.2

Este manual substitui o PDF da v1.0, que tinha divergências face ao programa (ver `CHANGELOG.md`). Descreve o comportamento real da versão 1.2.

## 1. Instalação e arranque

```
pip install -r requirements.txt
python TermiFolha.py
```

Requer Python 3.12 ou superior (testado em 3.13 e 3.15), com as versões mais recentes das dependências. Não instale o pacote antigo `fpdf` ao lado do `fpdf2`: partilham o mesmo nome de módulo. O programa funciona com ambos, mas só o `fpdf2` exporta PDF com caracteres fora de Latin-1.

Ao arrancar, o estado da sessão anterior (todas as folhas) é restaurado a partir de `.termifolha_estado.json`, que fica **na pasta do programa**. Erros de execução ficam em `.termifolha.log` (na mesma pasta). Para usar o programa em scripts, sem menus, veja a secção 12.

## 2. Conceitos

| Conceito | Descrição |
|---|---|
| Livro | O conjunto de folhas (separadores), como num ficheiro Excel. Fica todo guardado entre sessões. |
| Folha | Uma tabela. Só uma está activa de cada vez; os menus actuam sobre ela. |
| Coluna | Tem um nome único (sem distinguir maiúsculas) e um tipo fixo. |
| Linha | Identificada por um índice que começa em 0. |
| Tipos | texto, inteiro, decimal, data (`DD/MM/AAAA`). |

- Decimais usam **ponto** (`3.14`). A importação de CSV aceita também a vírgula decimal (`3,14`).
- Células em falta aparecem como `--`.
- Em qualquer prompt, escrever `sair` cancela a operação.
- Nos prompts com valor predefinido (`Enter = 5`), Enter aceita-o.
- O ecrã principal mostra no máximo 30 linhas; use *Análise básica → Ver por páginas* para ver o resto.

## 3. Menu principal

| Tecla | Acção |
|---|---|
| 1 | Folha: nova, importar, exportar, dummy, limpar, projectos |
| 2 | Linhas: adicionar, editar, remover, procurar e substituir |
| 3 | Colunas: adicionar, remover, renomear, formatar, fórmulas, tipo, regras, normalizar, faixas |
| 4 | Análise básica: ordenar, filtrar, início/fim, amostra, páginas, totais |
| 5 | Estatísticas (9 métodos) |
| 6 | Análise avançada: info, agrupar, duplicados, dados em falta, tabela dinâmica, outliers |
| 7 | Gráficos (10 tipos; opção de gravar em PNG) |
| 8 | Folhas: separadores, juntar, empilhar |
| 9 | Definições |
| S | Sobre |
| U / R | Desfazer / Refazer |
| Q | Sair |

## 4. Folha

| Opção | Descrição |
|---|---|
| [1] Nova folha | Pede nome e tipo de cada coluna; Enter num nome termina. |
| [2] Importar Excel | Lê `.xlsx`. Pede confirmação se já existir folha. Se uma coluna só tiver inteiros e células vazias, pergunta se deve ser inteira. **Um ficheiro com várias folhas importa todas** (uma por separador). |
| [3] Exportar Excel | Grava `.xlsx`. Se o livro tiver várias folhas, pergunta se quer gravá-las todas. |
| [4] Exportar Word | Colunas e tipos, tabela completa e estatísticas das colunas numéricas. |
| [5] Exportar PDF | O mesmo conteúdo. As células da tabela são truncadas a 14 caracteres. Com `fpdf2` usa a fonte DejaVu Sans (Unicode); com o `fpdf` antigo, caracteres fora de Latin-1 saem como `?`. |
| [6] Folha dummy | 8 linhas e 5 colunas de exemplo, com um valor em falta e uma linha duplicada. Pede confirmação se já existir folha. |
| [7] Limpar folha | Esvazia a folha activa (as outras folhas ficam). **Pode ser desfeita** com `[U]`. |
| [8] Importar CSV | Detecta o separador (`,` `;` tab `|`) e a codificação (UTF-8 ou Latin-1). Reconhece datas `DD/MM/AAAA` e `AAAA-MM-DD`, vírgula decimal e mantém códigos como `007` como texto. |
| [9] Exportar CSV | Formato internacional (`,` e `.`) ou Excel português (`;` e `,`). Grava em UTF-8 com BOM. |
| [A] Importar JSON | Uma lista de objectos (um por linha) ou um objecto de listas (uma por coluna). Inteiros com `null` ficam inteiros; estruturas aninhadas viram texto JSON. |
| [B] Colar dados | Cole texto copiado do Excel (tabulações) ou de um CSV. Termine com uma linha vazia ou `fim`. |
| [C] Exportar JSON | Lista de objectos; datas em `DD/MM/AAAA`, células em falta a `null`. Volta a importar sem perdas. |
| [D] Exportar Markdown | Tabela Markdown da folha activa, com os formatos de apresentação aplicados. |
| [E] Exportar HTML | Página HTML autónoma (sem scripts); todo o texto é escapado. |
| [P] Guardar projecto | Grava **todas as folhas**, fórmulas, regras e formatos num ficheiro `.termifolha`. |
| [O] Abrir projecto | Substitui o livro actual pelo de um ficheiro `.termifolha`. |

A extensão é acrescentada automaticamente se faltar. Pode indicar um caminho completo.

## 5. Linhas e colunas

**Linhas** — Adicionar: introduz linhas em ciclo, valor a valor, até escrever `sair`; os tipos e as regras de validação são verificados e as colunas calculadas não são pedidas. Editar célula: índice, coluna e novo valor (uma coluna calculada não se edita). Remover: mostra a linha e pede confirmação.

**Procurar e substituir** — Numa coluna ou em todas (Enter). O texto procurado e a substituição são **literais** (sem expressões regulares). Opções: só células inteiras, distinguir maiúsculas. Mostra quantas células mudam por coluna e pede confirmação. Nas colunas de texto substitui parte do texto; nas numéricas e de datas só substitui células com valor igual (a substituição vazia deixa a célula em branco). Se a substituição violar uma regra de validação, não altera nada.

**Colunas**

| Opção | Descrição |
|---|---|
| [1]–[3] | Adicionar (Enter deixa a célula em branco), remover e renomear. Os formatos, regras e visibilidade acompanham a coluna. |
| [4] Formatação de valores | Só apresentação: casas decimais (0–10), separador de milhares, estilo `1,234.50` ou `1 234,50`, prefixo e sufixo (`€ `, ` %`). Não altera os dados nem as exportações Excel/CSV/Word/PDF; aplica-se ao ecrã, ao Markdown e ao HTML. |
| [5] Colunas visíveis | Escolhe e ordena as colunas mostradas no ecrã (os dados não mudam). |
| [6] Coluna calculada | Fórmula sobre outras colunas; ver secção 10. |
| [7] Converter fórmula em valores | A coluna deixa de ser recalculada e passa a poder editar-se. |
| [8] Converter tipo da coluna | Texto ↔ inteiro ↔ decimal ↔ data. Diz quantos valores não convertem e só continua com a sua confirmação (esses ficam em branco). Um decimal com fracção **não** é arredondado para inteiro. Números e datas só se convertem passando por texto. |
| [9] Regras de validação | Por coluna: mínimo/máximo (números e datas), valores permitidos, obrigatório. Aplicam-se ao adicionar e editar e à substituição; "Verificar a folha" lista as violações existentes. |
| [A] Normalizar | Cria uma coluna nova: min-máx (0 a 1) ou z-score (média 0, desvio 1). A original não muda. |
| [B] Faixas | Cria uma coluna de texto com intervalos `(a, b]` (o primeiro inclui o mínimo): limites à escolha (`0,18,30,65`), N faixas de igual largura ou N faixas com igual número de linhas. |

## 6. Análise básica

- **Ordenar** por uma ou mais colunas, cada uma ascendente (`A`) ou descendente (`D`).
- **Filtrar** conforme o tipo da coluna: numérico (`== > < >= <=` e intervalo), data (antes, depois, entre) e texto (contém, igual, começa, termina; sem distinguir maiúsculas; o termo é literal). O resultado é só mostrado; a folha não muda.
- **Ver início/fim:** as primeiras ou últimas N linhas.
- **Amostra aleatória:** um número de linhas (`3`) ou uma percentagem (`20%`).
- **Ver por páginas:** `N` seguinte, `P` anterior, um número salta para a página, `sair` termina. O tamanho da página está nas Definições.
- **Linha de totais:** mostra sob a tabela a soma, média, contagem, mínimo ou máximo de cada coluna (calculado sobre todas as linhas). É só apresentação.

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
| [9] Correlação | Duas colunas: Pearson (relação linear) e Spearman (por ordens), com a leitura em palavras ("forte positiva"). Mais colunas: matriz de Pearson. Só usa as linhas onde todas têm valor. |

## 8. Análise avançada

- **Info:** dimensões, memória, tipo, valores não nulos e percentagem em falta por coluna.
- **Agrupar:** agrupa por uma coluna e agrega (soma, média, contagem, mínimo, máximo) uma coluna numérica.
- **Duplicados:** mostra e, se confirmar, remove mantendo a primeira ocorrência.
- **Dados em falta:** preencher (valor, média, mediana, *forward-fill*) ou remover linhas. Em colunas inteiras, média e mediana são **arredondadas**.
- **Tabela dinâmica:** linhas × colunas (opcional) × função (soma, média, contagem, mínimo, máximo) de uma coluna numérica, com totais opcionais. O resultado pode ser guardado como folha nova.
- **Valores anómalos (outliers):** pelo IQR (fora de `Q1 − k·IQR … Q3 + k·IQR`, k = 1,5) ou pelo desvio-padrão (a mais de k = 3 desvios da média). Mostra as linhas e deixa removê-las ou esvaziar a célula.

## 9. Gráficos

Cada gráfico abre numa janela do matplotlib; feche-a para voltar ao menu. Com **[B] Gravar gráficos em PNG** ligado, os gráficos não abrem janela: são gravados na pasta escolhida (por omissão `graficos/`) com o nome `<tipo>_<data>_<hora>.png`.

| | Gráfico | | Gráfico |
|---|---|---|---|
| 1 | Barras (eixos à escolha) | 6 | Histograma com quartis |
| 2 | Linhas (eixo X à escolha) | 7 | Boxplot |
| 3 | Circular | 8 | Heatmap de correlação |
| 4 | Dispersão | 9 | Linhas com média móvel |
| 5 | Área (soma acumulada) | A | Tendência polinomial |

## 10. Fórmulas (colunas calculadas)

*Colunas → [6] Coluna calculada.* Escreve-se o nome da nova coluna e a fórmula; a coluna é **recalculada automaticamente** sempre que a folha muda (editar, adicionar ou remover linhas, substituir, desfazer…).

| Elemento | Sintaxe |
|---|---|
| Colunas | `Salario` (sem distinguir maiúsculas) ou `[Preço unitário]` se tiver espaços |
| Números e texto | `12`, `3.5`, `"Lisboa"` |
| Operadores | `+ - * / // % **` |
| Comparações e lógica | `== != < <= > >=`, `and`, `or`, `not` |
| Funções | `abs`, `round(x, casas)`, `sqrt`, `log`, `exp`, `floor`, `ceil`, `min(a, b)`, `max(a, b)`, `se(condição, então, senão)` |
| Datas | `year(d)`, `month(d)`, `day(d)`, `days(d1, d2)` (diferença em dias) |

Exemplos: `Preco * Qtd`, `round(Salario / 12, 2)`, `se(Idade >= 18, 1, 0)`, `Nome + " (" + Cidade + ")"`, `year(Admissao)`.

- O tipo do resultado é inferido: inteiro (só com colunas inteiras, sem `/`, `sqrt`, `log`, `exp`), decimal, data ou texto. Uma comparação dá 1/0. Divisão por zero dá célula em branco.
- Uma célula em branco numa coluna usada propaga-se (o resultado fica em branco).
- Não se pode editar uma coluna calculada, nem renomear ou remover uma coluna que seja usada por uma fórmula: converta primeiro a fórmula em valores.
- **Segurança:** as fórmulas são interpretadas por um avaliador próprio com uma lista fechada de operações e funções; nunca são executadas como código Python. Uma fórmula adulterada num ficheiro de estado ou de projecto falha sem fazer nada.
- Limites: 500 caracteres e 200 elementos por fórmula.

## 11. Folhas (separadores)

*Menu principal → [8].* O cabeçalho mostra a folha activa quando há mais de uma.

| Opção | Descrição |
|---|---|
| [M] Mudar de folha | Por nome ou número. |
| [A] Adicionar | Folha nova com as colunas que definir. |
| [R] / [E] / [D] | Renomear, eliminar (nunca a última) e duplicar. Os nomes têm até 31 caracteres e não podem conter `[ ] : * ? /` nem `\`. |
| [J] Juntar tabelas | *Join* pela coluna chave de cada folha (a activa é a da esquerda). Mantém: só as linhas com correspondência, todas as da esquerda, todas as da direita ou todas. As colunas da direita com nome repetido recebem o sufixo `_<folha>`. Recusa chaves de tipos incompatíveis (texto com número) e junções que produzam mais de 1 000 000 de linhas. |
| [C] Empilhar | Põe as linhas de outra folha por baixo das da activa. As colunas unem-se pelo nome; as que faltam ficam em branco; inteiro com decimal dá decimal e outros tipos diferentes passam a texto (com aviso). |

Juntar e empilhar criam uma **folha nova** (`A+B`); as originais não mudam.

## 12. Linha de comandos (sem menus)

```
python TermiFolha.py --importar vendas.csv --exportar vendas.xlsx --exportar relatorio.html
python TermiFolha.py --projecto projecto.termifolha --exportar livro.xlsx --todas-as-folhas
python TermiFolha.py --importar dados.json --exportar dados.csv --csv-pt
```

| Argumento | Descrição |
|---|---|
| `--importar FICHEIRO` | `.xlsx` (todas as folhas), `.csv` ou `.json`. |
| `--projecto FICHEIRO` | Abre um `.termifolha`. |
| `--exportar FICHEIRO` | Formato pela extensão: `.xlsx .csv .json .md .html .docx .pdf`. Pode repetir-se. |
| `--csv-pt` | CSV em formato Excel português. |
| `--todas-as-folhas` | No `.xlsx`, grava um separador por folha. |
| `--versao`, `--help` | |

Neste modo o programa **não faz perguntas**, **não lê nem grava** o estado da sessão e devolve um código de saída: `0` sucesso, `1` erro (mensagem no `stderr`), `2` formato de exportação não suportado ou argumentos inválidos. Colunas só com inteiros e células vazias importadas de Excel ficam decimais (não há quem responda à pergunta).

## 13. Desfazer, refazer e definições

Cada alteração guarda um ponto de recuperação do **livro inteiro** (todas as folhas). `[U]` desfaz e `[R]` refaz. Por omissão guardam-se **20 passos**; o limite (1 a 100) muda-se em *Definições* (cada passo ocupa memória). Mudar de folha também se pode desfazer. Exporte periodicamente: o histórico não substitui uma cópia de segurança.

*Definições* também fixa o número de linhas por página (5–200). As definições ficam guardadas no estado, mas não nos projectos.

## 14. Resolução de problemas

| Problema | O que fazer |
|---|---|
| Erro ao importar Excel | Confirme a extensão `.xlsx`, que `openpyxl` está instalado e que o ficheiro não está aberto noutra aplicação. |
| Erro ao exportar | Confirme as permissões de escrita e que o ficheiro de destino não está aberto. Detalhes em `.termifolha.log`. |
| O programa avisa que o estado está corrompido | O ficheiro foi preservado como `.termifolha_estado.json.corrompido`; a sessão começa vazia. |
| Estado de uma versão antiga | Um `.termifolha_estado.pkl` (v1.0) é migrado automaticamente no primeiro arranque (cópia em `.pkl.migrado`). O estado da v1.1 é lido e convertido. **Depois de gravar com a v1.2, a v1.1 já não o consegue ler** (preserva-o como `.corrompido`). |
| "A fórmula de 'X' falhou" | A fórmula usa algo que já não existe (por exemplo, o ficheiro foi editado à mão). A coluna fica em branco; converta-a em valores ou recrie-a. |
| O gráfico não aparece | A janela pode ter aberto por trás de outras; ou a gravação em PNG está ligada (menu Gráficos, [B]). |
| Valores cortados no ecrã | Cada coluna é truncada a 15 caracteres no ecrã; os dados estão íntegros. |
| Faltam linhas no ecrã | O ecrã principal mostra 30; use *Ver por páginas*. |
| Folha com dados no arranque | É o estado da sessão anterior; use `[1]` → `[7]` para começar do zero. |
