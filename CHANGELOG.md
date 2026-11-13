# Changelog

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
