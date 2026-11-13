# TermiFolha

Folha de cálculo em linha de comando, em Python. Cria, edita, filtra e analisa tabelas com colunas tipadas (texto, inteiro, decimal, data), desenha gráficos e importa/exporta Excel, CSV, Word e PDF. O estado é gravado automaticamente entre sessões.

Autor: Gabriel Gomes Alves · Licença: MIT · Versão 1.0 · Manual completo: `Manual do Utilizador — TermiFolha v1.0.pdf`

## Instalação

Testado em Python 3.13. O código usa `zip(strict=True)`, por isso exige 3.10 ou superior; o pandas 3 exige 3.11 ou superior.

```
pip install -r requirements.txt
```

> `fpdf2` e o pacote antigo `fpdf` instalam ambos o módulo `fpdf` e **não devem coexistir**. Se o `pip list` mostrar `fpdf`, execute `pip uninstall fpdf` antes de instalar `fpdf2`. O programa funciona com ambos.

## Utilização

```
python TermiFolha.py
```

Navegue pelos menus com o número ou a letra da opção. Escrever `sair` num prompt cancela a operação. O manual descreve todos os menus.

Exemplo rápido: `[1] Folha` → `[6] Folha dummy` cria dados de exemplo; depois `[5] Estatísticas` → `[2]` → `Salario` mostra média, soma, mínimo, máximo e mediana.

### Novidades face ao manual v1.0

- **Folha → [8] Importar CSV / [9] Exportar CSV.** Detecta o separador (`,` `;` tab `|`) e a codificação (UTF-8 ou Latin-1). Datas `DD/MM/AAAA` são reconhecidas; códigos com zeros à esquerda (`007`) ficam como texto. A exportação usa UTF-8 com BOM para o Excel mostrar os acentos.
- **Amostra aleatória** aceita número de linhas (`3`) ou percentagem (`20%`).
- **Remover linha** mostra a linha antes de pedir confirmação.
- **Adicionar coluna:** Enter deixa a célula em branco (aparece `--`).

### Diferenças face ao manual (o código é a referência)

| Manual | Comportamento real |
|---|---|
| Folha dummy com 10 linhas | 8 linhas |
| Adicionar linhas "pede quantas" | introduz linhas em ciclo até escrever `sair` |
| Describe e Únicos "de todas as colunas" | pedem uma coluna |
| Limpar folha é irreversível e apaga o histórico | o passo entra no histórico e **pode ser desfeito** com `[U]` |
| Preencher com média/mediana | em colunas inteiras o valor é arredondado |

O ficheiro de estado `.termifolha_estado.pkl` fica **na pasta do programa**, não no directório onde se executa o comando.

## Desenvolvimento

```
pip install -r requirements-dev.txt
python -m pytest          # 53 testes
python -m ruff check .    # lint
```

Os testes simulam o teclado (fixture `app` em `tests/conftest.py`) e correm cada operação do programa. Avisos de depreciação (`DeprecationWarning`) falham os testes de propósito.

Estrutura:

```
TermiFolha.py        # aplicação (ponto de entrada)
tests/               # testes pytest
docs/                # MAPA-PROJECTO, DIAGNOSTICO, ARQUITECTURA, ROADMAP
requirements*.txt    # dependências
pyproject.toml       # configuração do pytest e do ruff
```

## Limitações conhecidas

- O estado é guardado com `pickle`: não abra ficheiros `.termifolha_estado.pkl` de origem desconhecida. Ver `docs/ROADMAP.md`.
- O manual PDF não foi actualizado e tem as divergências acima.
- Importar Excel não pede confirmação antes de substituir a folha actual (recuperável com `[U]`).
- PDF: caracteres fora de Latin-1 (ex.: `ż`) aparecem como `?`. Excel e Word não têm esta limitação.
