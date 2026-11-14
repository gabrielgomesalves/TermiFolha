# Manual do Utilizador — TermiFolha v1.4

Este manual descreve o comportamento real da versão 1.4. A versão em PDF é gerada a partir deste ficheiro (`python docs/gerar_manual_pdf.py`).

## 1. Instalação e arranque

```
pip install -r requirements.txt
python TermiFolha.py
```

Requer Python 3.12 ou superior (testado em 3.13 e 3.15), com as versões mais recentes das dependências. Para importar de bases de dados que não sejam SQLite (PostgreSQL, MySQL, …) instale também `pip install -r requirements-opcional.txt` e o driver da sua base. Não instale o pacote antigo `fpdf` ao lado do `fpdf2`: partilham o mesmo nome de módulo. O programa funciona com ambos, mas só o `fpdf2` exporta PDF com caracteres fora de Latin-1.

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
| 1 | Folha: nova, importar (Excel, CSV, JSON, URL, base de dados), exportar, dummy, limpar, projectos |
| 2 | Linhas: adicionar, editar, remover, procurar e substituir |
| 3 | Colunas: adicionar, remover, renomear, formatar, fórmulas, tipo, regras, normalizar, faixas |
| 4 | Análise básica: ordenar, filtrar, início/fim, amostra, páginas, totais |
| 5 | Estatísticas (9 métodos) |
| 6 | Análise avançada: info, agrupar, duplicados, dados em falta, tabela dinâmica, outliers |
| 7 | Gráficos (10 tipos; gravar em PNG; título e rótulos) |
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
| [F] Importar de URL | Descarrega um `.xlsx`, `.csv` ou `.json` de um endereço `http://` ou `https://` (até 50 MB, 30 s). Ver a secção 4.1. |
| [G] Importar de base de dados | Lê uma tabela, todas as tabelas ou o resultado de uma consulta `SELECT` de um ficheiro SQLite ou de outra base. Ver a secção 4.1. |
| [C] Exportar JSON | Lista de objectos; datas em `DD/MM/AAAA`, células em falta a `null`. Volta a importar sem perdas. |
| [D] Exportar Markdown | Tabela Markdown da folha activa, com os formatos de apresentação aplicados. |
| [E] Exportar HTML | Página HTML autónoma (sem scripts); todo o texto é escapado. |
| [P] Guardar projecto | Grava **todas as folhas**, fórmulas, regras e formatos num ficheiro `.termifolha`. |
| [O] Abrir projecto | Substitui o livro actual pelo de um ficheiro `.termifolha`. |

A extensão é acrescentada automaticamente se faltar. Pode indicar um caminho completo.

### 4.1. Importar de um endereço e de uma base de dados

**[F] Importar de URL.** O formato vem da extensão do endereço (`.xlsx`, `.csv`, `.tsv`, `.txt`, `.json`), do tipo de conteúdo que o servidor declara ou, em último caso, do início do ficheiro. O CSV e o JSON seguem as mesmas regras de [8] e [A]; um Excel com várias folhas importa todas. O programa só aceita `http://` e `https://`, recusa ficheiros com mais de 50 MB e desiste se o servidor não responder em 30 s. Pede confirmação se já existir folha.

- O endereço pode conter chaves de acesso (`?token=…`). O programa **nunca as mostra nem as regista**: nas mensagens e no registo aparece só `https://servidor/caminho`.
- Não há autenticação nem cabeçalhos próprios; para dados protegidos descarregue o ficheiro e use a importação normal.

**[G] Importar de base de dados.** Indique um ficheiro SQLite (`.db`, `.sqlite`, `.sqlite3`, ou `sqlite:///caminho`) ou, com o SQLAlchemy instalado, um endereço como `postgresql+psycopg://utilizador@servidor/base`. O programa mostra as tabelas e vistas e pergunta o que importar:

| Resposta | Resultado |
|---|---|
| Nome ou nº de uma tabela | A tabela inteira. |
| `*` | Todas as tabelas e vistas, **uma folha cada** (os nomes são ajustados às regras das folhas). Uma tabela vazia faz a importação falhar com um aviso. |
| Uma consulta `SELECT …` (ou `WITH …`) | O resultado da consulta. |

- **Só lê dados.** O SQLite abre-se em modo de leitura e a consulta tem de ser **um único** `SELECT`/`WITH`, sem `;` a meio nem palavras como `INSERT`, `UPDATE`, `DELETE`, `DROP` ou `PRAGMA`. Para servidores use mesmo assim uma conta só de leitura: o endereço é visível enquanto o escreve e a verificação do texto não substitui as permissões da base.
- Limite de 1 000 000 de linhas por consulta; acima disso o programa recusa e sugere `LIMIT`.
- Colunas repetidas numa consulta (`SELECT a.id, b.id`) ficam `id` e `id_2`; dados binários viram `<binário, N bytes>`; datas com fuso horário perdem o fuso. Texto `DD/MM/AAAA` ou `AAAA-MM-DD` passa a data e, tal como no Excel, o programa pergunta se uma coluna só com inteiros e vazios deve ser inteira.
- A palavra-passe do endereço nunca aparece nas mensagens de erro.

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

Com **[C] Título e rótulos dos gráficos** escolhe o título e os rótulos dos eixos X e Y de **um tipo de gráfico** ou de **todos** (`[0]`). Os textos **ficam gravados com a folha** (no estado, nos projectos e no desfazer) e usam-se sempre que o gráfico é desenhado, no ecrã ou em PNG, até os alterar. Nos prompts, Enter mantém o que lá está e `-` apaga esse texto (volta ao automático); `[L]` apaga todos os textos da folha. Um texto de um gráfico específico sobrepõe-se ao de «todos». O gráfico circular não tem eixos: usa só o título. O menu mostra quantos textos a folha tem guardados. O texto é limitado a 120 caracteres e o `$` é escrito tal como está (não é lido como fórmula). Nos gráficos sem eixos (circular) só o título tem efeito visível.

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
| Células | `B2`: coluna **B** (a 2.ª) e linha **2**, como na coluna `#` do ecrã (a primeira linha é a 0). `$B$2` é igual. Também `[Salario]2` |
| Intervalos | `B2:B9` (ou `B0:C7`, que abrange as duas colunas); só dentro de `soma`, `media`, `contar`, `min`, `max` |
| Totais | `soma(x)`, `media(x)`, `contar(x)`, `min(x)`, `max(x)` com **um** argumento (coluna, intervalo ou expressão); `min(a, b)` e `max(a, b)` continuam a comparar linha a linha |
| Outras linhas | `anterior(x)`, `seguinte(x)`, `desloca(x, n)`: o valor de `x` noutra linha, **relativo à linha actual** (ver abaixo) |
| Outras folhas | `Vendas!B2`, `'Vendas 2025'!B2:B9`, `Vendas![Preço]3` e, só dentro de totais, a coluna inteira `Vendas![Preço]` (ver abaixo) |

Exemplos: `Preco * Qtd`, `round(Salario / 12, 2)`, `se(Idade >= 18, 1, 0)`, `Nome + " (" + Cidade + ")"`, `year(Admissao)`, `Salario / soma(Salario) * 100`, `Salario / $C$0`, `soma(C0:C7)`, `Total - anterior(Total)`, `Preco * (1 + Taxas!A0)`.

**Células e intervalos**

- As letras seguem a ordem das colunas (`A` = 1.ª, `B` = 2.ª, … `AA` = 27.ª); o ecrã da fórmula mostra a correspondência (`A=Nome  B=Idade …`). A linha é o número da coluna `#`.
- Uma referência de célula (`B2`) é sempre a **mesma célula** em todas as linhas da coluna calculada (`Salario / $C$0` divide tudo pelo primeiro salário). Para usar o valor da própria linha, escreva o nome da coluna; para usar o de outra linha, em relação à actual, use `anterior`, `seguinte` ou `desloca` (abaixo).
- Ao gravar, as letras são trocadas pelos nomes das colunas (`B2` passa a `[Idade]2`), por isso a fórmula continua certa se as colunas mudarem de posição. O ecrã mostra "Guardada como: …".
- Uma coluna referida por uma célula ou intervalo conta como usada: não se remove, renomeia nem converte.
- Um intervalo abrange todas as colunas entre os extremos. Vazios são ignorados por `soma`, `media`, `min` e `max`; `contar` conta as células com valor, de qualquer tipo. `soma` e `media` não aceitam datas; `min` e `max` aceitam.
- Uma linha que não exista (por exemplo, depois de a remover), uma folha que já não exista ou uma coluna em falta fazem a fórmula falhar com um aviso, e a coluna fica em branco.
- Se uma coluna se chamar como uma célula (por exemplo `Q1`), o nome da coluna ganha.

**Referências relativas (que acompanham a linha)**

Uma fórmula vale para a coluna inteira, por isso o que "anda" com a linha é feito por funções que deslocam uma coluna ou uma expressão:

| Função | Valor em cada linha |
|---|---|
| `anterior(x)` | `x` na linha de cima (a primeira linha fica em branco). `anterior(x, 3)`: três linhas acima. |
| `seguinte(x)` | `x` na linha de baixo (a última fica em branco). `seguinte(x, 2)`: duas linhas abaixo. |
| `desloca(x, n)` | `x` na linha `n` posições à frente (`n` negativo = atrás). `desloca(x, -1)` é `anterior(x)`. |

- O deslocamento tem de ser um **número inteiro fixo** (não pode vir de uma coluna). As linhas que ficam sem origem ficam em branco; um inteiro com células em branco continua inteiro.
- Servem para qualquer tipo (números, texto, datas): `Salario - anterior(Salario)` (variação), `se(Salario > anterior(Salario), 1, 0)`, `days(Data, anterior(Data))`.
- A ordem das linhas é a do ecrã; se ordenar a folha, a coluna recalcula-se com a nova ordem.

**Referências a outras folhas**

- `Vendas!B2` é a célula da coluna B (a 2.ª) da folha *Vendas*, linha 2; `Vendas!B0:B9` é um intervalo; `Vendas![Preço]3` usa o nome da coluna. As letras e os nomes são os **da outra folha**.
- Se o nome da folha tiver espaços ou símbolos, ponha-o entre aspas: `'Vendas 2025'!B2` ou `"Vendas 2025"!B2`. Não se distingue maiúsculas de minúsculas.
- A coluna inteira de outra folha (`Vendas![Preço]`) só pode estar dentro de `soma`, `media`, `contar`, `min` ou `max`, porque as folhas têm números de linhas diferentes: `Salario / soma(Vendas![Preço])`. Uma célula ou um intervalo de outra folha usa-se como qualquer outro valor.
- Ao gravar, a fórmula fica em forma canónica, com o nome da folha entre aspas e as colunas por nome: `'Vendas 2025'![Preço]2`.
- **Recálculo:** todas as colunas calculadas do livro são recalculadas a cada alteração, **pela ordem das dependências** (uma coluna calculada de outra folha é calculada antes de quem a usa). Se duas fórmulas dependerem uma da outra (só possível num ficheiro editado à mão), ficam em branco com o aviso «referência circular».
- Uma folha usada pela fórmula de outra **não se elimina** (converta primeiro essas colunas em valores) e uma coluna usada por outra folha não se remove, renomeia nem converte; o aviso diz qual (`Folha!coluna`). **Renomear** uma folha actualiza as fórmulas que a referem.
- Os nomes de folha não podem ter aspas simples e duplas ao mesmo tempo.

- O tipo do resultado é inferido: inteiro (só com colunas inteiras, sem `/`, `sqrt`, `log`, `exp`, `media`), decimal, data ou texto. Uma comparação dá 1/0. Divisão por zero dá célula em branco.
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
| [R] / [E] / [D] | Renomear, eliminar (nunca a última) e duplicar. Os nomes têm até 31 caracteres, não podem conter `[ ] : * ? /` nem `\` e não podem ter aspas simples e duplas ao mesmo tempo. Uma folha usada pela fórmula de outra não se elimina (secção 10). |
| [J] Juntar tabelas | *Join* pela coluna chave de cada folha (a activa é a da esquerda). Mantém: só as linhas com correspondência, todas as da esquerda, todas as da direita ou todas. As colunas da direita com nome repetido recebem o sufixo `_<folha>`. Recusa chaves de tipos incompatíveis (texto com número) e junções que produzam mais de 1 000 000 de linhas. |
| [C] Empilhar | Põe as linhas de outra folha por baixo das da activa. As colunas unem-se pelo nome; as que faltam ficam em branco; inteiro com decimal dá decimal e outros tipos diferentes passam a texto (com aviso). |

Juntar e empilhar criam uma **folha nova** (`A+B`); as originais não mudam.

## 12. Linha de comandos (sem menus)

```
python TermiFolha.py --importar vendas.csv --exportar vendas.xlsx --exportar relatorio.html
python TermiFolha.py --projecto projecto.termifolha --exportar livro.xlsx --todas-as-folhas
python TermiFolha.py --importar dados.json --exportar dados.csv --csv-pt
python TermiFolha.py --importar https://exemplo.pt/dados.csv --exportar dados.xlsx
python TermiFolha.py --base-dados loja.db --tabela clientes --exportar clientes.csv
python TermiFolha.py --base-dados loja.db --consulta "SELECT nome, saldo FROM clientes WHERE saldo > 0" --exportar saldos.xlsx
python TermiFolha.py --base-dados loja.db --tabela "*" --exportar loja.xlsx --todas-as-folhas
```

| Argumento | Descrição |
|---|---|
| `--importar FICHEIRO_OU_URL` | `.xlsx` (todas as folhas), `.csv` ou `.json`, num ficheiro ou num endereço `http(s)://`. |
| `--base-dados FICHEIRO_OU_ENDEREÇO` | Importa de um ficheiro SQLite ou de um endereço SQLAlchemy. Exige `--tabela` ou `--consulta`. Não se combina com `--importar`. |
| `--tabela NOME` | Com `--base-dados`: a tabela a importar; `*` importa todas (uma folha cada). |
| `--consulta SELECT` | Com `--base-dados`: importa o resultado de uma consulta de leitura. |
| `--projecto FICHEIRO` | Abre um `.termifolha`. |
| `--exportar FICHEIRO` | Formato pela extensão: `.xlsx .csv .json .md .html .docx .pdf`. Pode repetir-se. |
| `--csv-pt` | CSV em formato Excel português. |
| `--todas-as-folhas` | No `.xlsx`, grava um separador por folha. |
| `--versao`, `--help` | |

Neste modo o programa **não faz perguntas**, **não lê nem grava** o estado da sessão e devolve um código de saída: `0` sucesso, `1` erro (mensagem no `stderr`), `2` formato de exportação não suportado ou argumentos inválidos (por exemplo `--base-dados` sem `--tabela` nem `--consulta`). Colunas só com inteiros e células vazias importadas de Excel ficam decimais (não há quem responda à pergunta).

## 13. Desfazer, refazer e definições

Cada alteração guarda um ponto de recuperação do **livro inteiro** (todas as folhas). `[U]` desfaz e `[R]` refaz. Por omissão guardam-se **20 passos**; o limite (1 a 100) muda-se em *Definições* (cada passo ocupa memória). Mudar de folha também se pode desfazer. Exporte periodicamente: o histórico não substitui uma cópia de segurança.

*Definições* também fixa o número de linhas por página (5–200). As definições ficam guardadas no estado, mas não nos projectos.

## 14. Resolução de problemas

| Problema | O que fazer |
|---|---|
| Erro ao importar Excel | Confirme a extensão `.xlsx`, que `openpyxl` está instalado e que o ficheiro não está aberto noutra aplicação. |
| Erro ao importar de URL | O endereço tem de ser `http://` ou `https://`, o servidor tem de responder em 30 s e o ficheiro não pode passar os 50 MB. O programa diz se o servidor respondeu com erro (por exemplo 404) ou se não foi possível ligar. |
| Erro ao importar de base de dados | Confirme o caminho do ficheiro SQLite, o nome da tabela e que a consulta é um só `SELECT`. Para outras bases instale o SQLAlchemy e o driver (`requirements-opcional.txt`). |
| Erro ao exportar | Confirme as permissões de escrita e que o ficheiro de destino não está aberto. Detalhes em `.termifolha.log`. |
| O programa avisa que o estado está corrompido | O ficheiro foi preservado como `.termifolha_estado.json.corrompido`; a sessão começa vazia. |
| Estado de uma versão antiga | Um `.termifolha_estado.pkl` (v1.0) é migrado automaticamente no primeiro arranque (cópia em `.pkl.migrado`). O estado da v1.1 é lido e convertido. **Depois de gravar com a v1.2 ou posterior, a v1.1 já não o consegue ler** (preserva-o como `.corrompido`). |
| "A fórmula de 'X' falhou" | A fórmula usa algo que já não existe (uma coluna, uma linha ou uma folha) ou o ficheiro foi editado à mão. A coluna fica em branco; converta-a em valores ou recrie-a. |
| O gráfico não aparece | A janela pode ter aberto por trás de outras; ou a gravação em PNG está ligada (menu Gráficos, [B]). |
| Valores cortados no ecrã | Cada coluna é truncada a 15 caracteres no ecrã; os dados estão íntegros. |
| Faltam linhas no ecrã | O ecrã principal mostra 30; use *Ver por páginas*. |
| Folha com dados no arranque | É o estado da sessão anterior; use `[1]` → `[7]` para começar do zero. |
