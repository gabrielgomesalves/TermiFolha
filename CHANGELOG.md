# Changelog

## [Não lançado]

### Adicionado
- Importar e exportar CSV (menu Folha `[8]` e `[9]`).
- Amostra aleatória por percentagem (`20%`).
- Remover linha mostra a linha antes de confirmar.
- Suite de 53 testes (`pytest`), `requirements.txt`, `requirements-dev.txt`, configuração do `ruff`, `.gitignore`, `.gitattributes`, documentação em `docs/`.

### Corrigido
- **Alto:** preencher valores em falta com média ou mediana numa coluna inteira fazia o programa terminar com `TypeError`.
- **Alto:** colunas de data criadas em "Nova folha" ficavam como texto, por isso o filtro e as estatísticas de data não funcionavam.
- **Alto:** importar Excel com cabeçalhos numéricos (anos) fazia várias operações rebentar.
- **Médio:** erros a gravar ou a ler o estado eram engolidos; um ficheiro corrompido era sobrescrito. Agora o utilizador é avisado e o ficheiro é preservado como `.corrompido`.
- **Médio:** o ficheiro de estado seguia o directório de trabalho; agora fica na pasta do programa.
- **Médio:** adicionar coluna não aceitava células em branco, ao contrário do manual.
- **Médio:** filtro de texto com operador inválido filtrava como "contém" sem avisar; o termo era interpretado como expressão regular (`(` provocava erro).
- **Médio:** exportar PDF abortava com caracteres fora de Latin-1 e usava API deprecada do `fpdf2` (fonte `arial`, `ln=True`). Funciona agora com `fpdf` e `fpdf2`.
- Removido `inplace=True` (deprecado no pandas 3.1).

### Alterado
- Importar Excel passa a só alterar a folha activa depois de a importação terminar com sucesso.
