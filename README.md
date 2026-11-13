# TermiFolha

Folha de cálculo em linha de comando, em Python, num único ficheiro (`TermiFolha.py`). Cria, edita, filtra e analisa tabelas com colunas tipadas (texto, inteiro, decimal, data), desenha gráficos e importa/exporta Excel e CSV, e exporta Word e PDF. O estado é gravado automaticamente entre sessões.

Autor: Gabriel Gomes Alves · Licença: MIT · Versão 1.1 · Manual: [`docs/MANUAL.md`](docs/MANUAL.md)

## Instalação

```
pip install -r requirements.txt
python TermiFolha.py
```

Testado em Python 3.13. O pandas 3 exige 3.11 ou superior.

> `fpdf2` e o pacote antigo `fpdf` instalam ambos o módulo `fpdf` e **não devem coexistir**. Se o `pip list` mostrar `fpdf`, execute `pip uninstall fpdf` antes de instalar `fpdf2`. O programa funciona com ambos, mas só o `fpdf2` exporta PDF com caracteres fora de Latin-1 (`ż`, `€`, …).

## Utilização

Navegue pelos menus com o número ou a letra da opção; `sair` cancela uma operação. O manual descreve todos os menus.

Exemplo: `[1] Folha` → `[6] Folha dummy` cria dados de exemplo; depois `[5] Estatísticas` → `[2]` → `Salario` mostra média, soma, mínimo, máximo e mediana.

O estado fica em `.termifolha_estado.json` e os erros em `.termifolha.log`, ambos na pasta do programa.

### Ficheiros gerados e actualização da v1.0

A v1.0 guardava o estado em `.termifolha_estado.pkl`. Ao arrancar, a v1.1 migra-o automaticamente (com uma leitura restrita que só aceita DataFrames) e guarda uma cópia como `.pkl.migrado`.

## Desenvolvimento

```
pip install -r requirements-dev.txt
python -m pytest          # 90 testes
python -m ruff check .    # lint
```

Os testes simulam o teclado (fixture `app` em `tests/conftest.py`). Avisos de depreciação (`DeprecationWarning`) falham os testes de propósito. Dois testes dependem da versão do `fpdf` e são saltados no ambiente a que não se aplicam.

```
TermiFolha.py        # toda a aplicação (ponto de entrada)
tests/               # testes pytest
docs/                # MANUAL, ARQUITECTURA, MAPA-PROJECTO, DIAGNOSTICO, ROADMAP, RENOMEACOES
requirements*.txt    # dependências
pyproject.toml       # configuração do pytest e do ruff
```

O código está em inglês (funções, variáveis, constantes); mensagens, comentários e documentação em português. A equivalência com os nomes da v1.0 está em [`docs/RENOMEACOES.md`](docs/RENOMEACOES.md).

## Limitações conhecidas

- O manual em PDF da v1.0 está desactualizado; o de referência é `docs/MANUAL.md`.
- A tabela do PDF trunca as células a 14 caracteres.
- Com o `fpdf` antigo, caracteres fora de Latin-1 aparecem como `?` no PDF.
- O histórico de desfazer tem 4 passos.
- Os nomes das funções de teste estão em português (descrevem o comportamento esperado).
