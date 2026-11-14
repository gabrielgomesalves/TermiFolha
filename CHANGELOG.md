# Changelog

## [1.4.0]

### Adicionado
- **Referências relativas nas fórmulas** que acompanham a linha: `anterior(x)`, `seguinte(x)` e `desloca(x, n)` (por exemplo `Salario - anterior(Salario)`), para números, texto e datas.
- **Referências a outras folhas** nas fórmulas: `Vendas!B2`, `'Vendas 2025'!B2:B9`, `Vendas![Preço]3` e, dentro de `soma`/`media`/`contar`/`min`/`max`, a coluna inteira `Vendas![Preço]`.
- **Recálculo de todo o livro** por ordem de dependências: uma coluna calculada de outra folha é calculada antes de quem a usa; ciclos são detectados e avisados («referência circular»).
- **Título e rótulos dos gráficos guardados** com a folha (estado, projectos e desfazer), por tipo de gráfico ou para todos; `-` apaga um texto e `[L]` apaga todos.
- **Importar de URL** (`[F]` no menu Folha): `.xlsx`, `.csv` e `.json` em `http://` e `https://`, com limite de 50 MB e de 30 s; as chaves no endereço nunca são mostradas nem registadas.
- **Importar de bases de dados** (`[G]`): tabela, todas as tabelas (uma folha cada) ou consulta `SELECT` de um ficheiro SQLite; outras bases (PostgreSQL, MySQL, …) com o SQLAlchemy opcional (`requirements-opcional.txt`). Só leitura, com limite de 1 000 000 de linhas.
- Linha de comandos: `--importar` aceita endereços `http(s)://`; novos `--base-dados`, `--tabela` e `--consulta`.
- 113 novos testes (611 no total), incluindo um servidor HTTP local e bases SQLite criadas na hora.

### Alterado
- O título e os rótulos dos gráficos já **não se gastam** ao primeiro gráfico: valem até os alterar. O menu `[C]` passou a escolher o gráfico.
- As colunas calculadas de **todas** as folhas são recalculadas a cada alteração (antes só as da folha activa).
- Uma folha usada pela fórmula de outra não se elimina; uma coluna usada por fórmulas de outra folha não se remove, renomeia nem converte; renomear uma folha actualiza as fórmulas que a referem.
- Os nomes de folha não podem ter aspas simples e duplas ao mesmo tempo.
- O estado ganha a chave `graficos` por folha (o formato continua na versão 2; estados antigos abrem sem ela). Versão 1.4.

### Interno
- `evaluate_formula`, `canonical_formula` e `_parse_formula` recebem opcionalmente o livro (`book`); `recalculate_formulas` trabalha sobre o livro inteiro.

## [1.3.0]

### Adicionado
- **Fórmulas com referências a células e intervalos**: `B2`, `$B$2`, `[Salario]2`, `B2:B9`; letras de coluna (A, B, … AA) mostradas no ecrã da fórmula; linhas como na coluna `#`.
- **Agregados em fórmulas**: `soma`, `media`, `contar`, `min(x)`, `max(x)` sobre uma coluna, um intervalo ou uma expressão (`Salario / soma(Salario) * 100`).
- **Título e rótulos dos eixos configuráveis** para o próximo gráfico (Gráficos → `[C]`), no ecrã ou em PNG.
- **Manual em PDF regenerado** (v1.3) a partir de `docs/MANUAL.md`, com o gerador `docs/gerar_manual_pdf.py`; substitui o PDF da v1.0.

### Alterado
- As fórmulas são gravadas com os nomes das colunas (`B2` → `[Idade]2`), para não dependerem da posição.
- `min` e `max` aceitam um argumento (agregado) ou dois (linha a linha).
- O texto entre aspas numa fórmula nunca é lido como nome de coluna ou célula.
- Versão 1.3. O formato do estado não mudou (v2).

## [1.2.0]

### Adicionado
- **Várias folhas (separadores)** num livro: mudar, adicionar, renomear, eliminar, duplicar; **juntar** (join: interna, esquerda, direita, completa) e **empilhar** (concatenar) criam folha nova e preservam os tipos. O Excel com várias folhas importa todas e pode exportá-las todas.
- **Colunas calculadas por fórmula** (`Preco * Qtd`, `se(Idade >= 18, 1, 0)`, `year(Admissao)`…), recalculadas automaticamente. Avaliador próprio sobre a árvore sintáctica (nunca `eval`/`exec`).
- **Procurar e substituir** (literal, com pré-visualização, respeita as regras), **converter tipo de coluna** (reporta o que não converte; não arredonda inteiros em silêncio) e **regras de validação** por coluna (mínimo/máximo, permitidos, obrigatório).
- Apresentação: **formatação de valores** (casas decimais, milhares, `€`), **colunas visíveis** (escolher e ordenar), **ver por páginas**, **linha de totais**; o ecrã principal mostra no máximo 30 linhas.
- Análise: **correlação** (Pearson e Spearman, matriz), **valores anómalos** (IQR e desvio-padrão), **normalizar** (min-máx, z-score), **faixas** (limites à escolha, largura igual, quantis) e **tabela dinâmica**.
- **Gráficos em PNG**: os 10 gráficos podem ser gravados em ficheiro em vez de abrirem uma janela.
- **JSON** (importar/exportar), **colar dados** do Excel ou de CSV, exportar **Markdown** e **HTML**, e **ficheiros de projecto** `.termifolha` com todas as folhas.
- **Modo de linha de comandos** (`--importar`, `--projecto`, `--exportar`, `--csv-pt`, `--todas-as-folhas`) sem perguntas e sem tocar no estado da sessão; códigos de saída 0/1/2.
- Menu **Definições**: passos de desfazer (1–100) e linhas por página.
- Datas `AAAA-MM-DD` reconhecidas na importação de CSV/JSON/texto colado.
- Mais de 380 novos testes (471 no total), incluindo um que garante que cada opção de cada menu chama a função certa.

### Alterado
- O histórico de desfazer guarda o **livro inteiro** e passa de 3 para **20 passos** por omissão (a v1.1 guardava 4 estados, ou seja, 3 passos).
- O estado em JSON passa à **versão 2** (várias folhas e metadados). A versão 1 continua a ser lida; **a v1.1 não consegue ler a versão 2**.
- *Limpar folha* só esvazia a folha activa quando há várias.
- `export_word`/`export_pdf`/`export_csv` passam a usar funções `write_*` reutilizáveis (também usadas pela linha de comandos).
- Versão 1.2.

## [1.1.0]

### Adicionado
- Importar e exportar CSV (menu Folha `[8]` e `[9]`): separador e codificação detectados, datas `DD/MM/AAAA`, vírgula decimal, códigos com zeros à esquerda preservados; exportação em formato internacional ou Excel português.
- Amostra aleatória por percentagem (`20%`).
- Describe e Únicos aceitam Enter = todas as colunas.
- Registo de erros em `.termifolha.log` (rotativo).
- Importar Excel/CSV pede confirmação se já existir folha; o Excel pergunta se colunas só com inteiros e células vazias devem ser inteiras.
- Suite de 90 testes (`pytest`), `requirements*.txt`, configuração do `ruff`, CI, `.gitignore`, `.gitattributes`, `LICENSE`, e documentação em `docs/` (incluindo o novo `MANUAL.md`).

### Alterado
- **O estado passa a ser gravado em JSON** (`.termifolha_estado.json`) com escrita atómica. Um `.termifolha_estado.pkl` da v1.0 é migrado automaticamente no primeiro arranque, com uma leitura restrita que só aceita DataFrames, e fica guardado como `.pkl.migrado`.
- Identificadores do código em inglês (`docs/RENOMEACOES.md`). Mensagens ao utilizador mantêm-se em português.
- Remover linha mostra a linha antes de confirmar.
- Importar passa a só alterar a folha activa depois de a importação terminar com sucesso.
- O PDF usa a fonte DejaVu Sans com `fpdf2` e repõe os acentos nos títulos.
- Versão 1.1.
- Dependências sempre nas versões mais recentes (`pandas>=3.1`, `numpy>=2.5`, `matplotlib>=3.11`, `fpdf2>=2.9`, …). Em consequência, o Python mínimo passa a 3.12.
- Licença: removido do cabeçalho o "todos os direitos reservados", que contradizia a MIT.

### Corrigido
- **Alto:** preencher valores em falta com média ou mediana numa coluna inteira fazia o programa terminar com `TypeError` (agora arredonda).
- **Alto:** colunas de data criadas em "Nova folha" ficavam como texto, por isso o filtro e as estatísticas de data não funcionavam.
- **Alto:** importar Excel com cabeçalhos numéricos (anos) fazia várias operações rebentar.
- **Segurança:** o estado deixou de ser `pickle`, que executa código ao carregar um ficheiro adulterado.
- **Médio:** erros a gravar ou a ler o estado eram engolidos; um ficheiro corrompido era sobrescrito. Agora o utilizador é avisado e o ficheiro é preservado como `.corrompido`.
- **Médio:** o ficheiro de estado seguia o directório de trabalho; agora fica na pasta do programa.
- **Médio:** adicionar coluna não aceitava células em branco.
- **Médio:** o filtro de texto interpretava o termo como expressão regular (`(` provocava erro) e aceitava operadores inválidos em silêncio.
- **Médio:** exportar PDF abortava com caracteres fora de Latin-1 e usava API deprecada do `fpdf2`.
- Removido `inplace=True` (deprecado no pandas 3.1).
