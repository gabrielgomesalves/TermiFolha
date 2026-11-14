<div align="center">

# TermiFolha

**Uma folha de cálculo completa, dentro do terminal.**

Tabelas tipadas, várias folhas, fórmulas, análise estatística, gráficos e exportação para Excel, Word, PDF e mais — num único ficheiro Python.

![Versão](https://img.shields.io/badge/versão-1.3-blue)
![Python](https://img.shields.io/badge/python-3.12%2B-3776AB?logo=python&logoColor=white)
![Licença](https://img.shields.io/badge/licença-MIT-green)
![Testes](https://img.shields.io/badge/testes-498-brightgreen)

[Manual](docs/MANUAL.md) · [Manual em PDF](<Manual do Utilizador — TermiFolha v1.3.pdf>) · [Changelog](CHANGELOG.md) · [Arquitectura](docs/ARQUITECTURA.md) · [Roadmap](docs/ROADMAP.md)

</div>

---

## Porquê o TermiFolha?

- **Sem interface gráfica, sem fricção.** Corre em qualquer terminal, navega-se por menus com uma tecla e funciona em máquinas onde não há Excel.
- **Um único ficheiro.** Toda a aplicação está em `TermiFolha.py`: copie-o, leia-o, modifique-o.
- **Dados com tipos a sério.** Texto, inteiro, decimal e data, com regras de validação por coluna.
- **Seguro por desenho.** As fórmulas são avaliadas por um interpretador próprio, que nunca usa `eval` nem `exec`.
- **O trabalho nunca se perde.** O estado é gravado entre sessões e há desfazer/refazer.
- **Também em scripts.** O modo de linha de comandos converte e exporta ficheiros sem fazer perguntas.

## Funcionalidades

| | |
|---|---|
| **Folhas e livros** | Várias folhas por livro; juntar (*join*) e empilhar; ficheiros de projecto `.termifolha` |
| **Fórmulas** | `Preco * Qtd`, `se(Idade >= 18, 1, 0)`, `Salario / soma(Salario) * 100`, referências `B2` e intervalos `B2:B9` |
| **Edição** | Linhas e colunas, procurar e substituir com pré-visualização, conversão de tipos, colar dados do Excel |
| **Limpeza** | Valores em falta, duplicados, valores anómalos (IQR e desvio-padrão), normalização, faixas |
| **Análise** | Filtros, ordenação, agrupamento, tabelas dinâmicas, correlação (Pearson e Spearman), estatísticas |
| **Gráficos** | 10 tipos, com título e rótulos à escolha, no ecrã ou gravados em PNG |
| **Importar** | Excel (todas as folhas), CSV (detecta separador e codificação), JSON |
| **Exportar** | Excel, CSV, JSON, Markdown, HTML, Word e PDF |

## Começar

Requer **Python 3.12 ou superior**.

```
git clone <endereço-do-repositório>
cd TermiFolha
pip install -r requirements.txt
python TermiFolha.py
```

### Primeiros passos

1. No menu principal, `[1] Folha` → `[6] Folha dummy` cria dados de exemplo.
2. `[5] Estatísticas` → `[2]` → `Salario` mostra a média, a soma, o mínimo, o máximo e a mediana.
3. `[3] Colunas` permite acrescentar uma coluna calculada, por exemplo `Salario * 12`.
4. `[7] Gráficos` desenha os dados; `[1] Folha` exporta o resultado.

Navegue com o número ou a letra de cada opção. Escreva `sair` para cancelar uma operação e `U` / `R` para desfazer e refazer.

## Em scripts

Sem menus, sem perguntas e sem tocar no estado da sessão:

```
python TermiFolha.py --importar vendas.csv --exportar vendas.xlsx --exportar relatorio.html
python TermiFolha.py --projecto projecto.termifolha --exportar livro.xlsx --todas-as-folhas
```

Formatos de exportação: `.xlsx .csv .json .md .html .docx .pdf`. O código de saída é `0` (sucesso), `1` (erro) ou `2` (formato não suportado). Todas as opções estão na secção 12 do [manual](docs/MANUAL.md).

## Ficheiros criados

| Ficheiro | Conteúdo |
|---|---|
| `.termifolha_estado.json` | Estado da sessão (todas as folhas), na pasta do programa |
| `.termifolha.log` | Registo de erros |

Quem vem da v1.0 não perde nada: o estado antigo (`.termifolha_estado.pkl`) é migrado automaticamente no primeiro arranque, e fica uma cópia em `.pkl.migrado`. Depois de gravar com a v1.2 ou posterior, a v1.1 já não lê o estado.

> **Nota sobre PDF.** O `fpdf2` e o pacote antigo `fpdf` instalam o mesmo módulo e não devem coexistir. Se o `pip list` mostrar `fpdf`, execute `pip uninstall fpdf` antes de instalar `fpdf2`. Só o `fpdf2` exporta caracteres fora de Latin-1 (`ż`, `€`, …).

## Documentação

| Documento | Para quê |
|---|---|
| [Manual do Utilizador](docs/MANUAL.md) | Todos os menus, fórmulas e opções (também em [PDF](<Manual do Utilizador — TermiFolha v1.3.pdf>)) |
| [Changelog](CHANGELOG.md) | O que mudou em cada versão |
| [Arquitectura](docs/ARQUITECTURA.md) | Como o código está organizado por dentro |
| [Roadmap](docs/ROADMAP.md) | O que falta fazer |
| [Renomeações](docs/RENOMEACOES.md) | Equivalência entre os nomes da v1.0 e os actuais |

## Desenvolvimento

O projecto usa um ambiente próprio em `.venv/` (ignorado pelo git).

```
python -m venv .venv
.venv\Scripts\python.exe -m pip install -U -r requirements-dev.txt   # Windows
.venv/bin/python -m pip install -U -r requirements-dev.txt           # Linux e macOS

python -m pytest          # testes
python -m ruff check .    # lint
```

Os testes simulam o teclado (fixture `app` em `tests/conftest.py`). Avisos de depreciação (`DeprecationWarning`) falham os testes de propósito, para apanhar quebras cedo. Dois testes dependem da versão do `fpdf` e são saltados no ambiente a que não se aplicam. A integração contínua (`.github/workflows/ci.yml`) corre o lint e os testes a cada alteração.

```
TermiFolha.py        # toda a aplicação (ponto de entrada)
tests/               # testes pytest
docs/                # manual, arquitectura, roadmap, ...
requirements*.txt    # dependências
pyproject.toml       # configuração do pytest e do ruff
```

O código está em inglês (funções, variáveis, constantes); as mensagens, os comentários e a documentação estão em português de Portugal.

## Limitações conhecidas

- A tabela do PDF trunca as células a 14 caracteres.
- Com o `fpdf` antigo, os caracteres fora de Latin-1 aparecem como `?` no PDF.
- O histórico de desfazer guarda o livro inteiro: 20 passos por omissão (configurável até 100), e cada passo usa memória.
- As referências a células (`B2`) são sempre fixas: não acompanham a linha e não funcionam entre folhas.
- O título e os rótulos dos gráficos valem só para o gráfico seguinte e não ficam gravados.
- O manual em PDF é gerado de `docs/MANUAL.md` (`python docs/gerar_manual_pdf.py`); em caso de dúvida vale o `.md`.

## Licença e autor

Distribuído sob a licença [MIT](LICENSE). Criado por Gabriel Gomes Alves.
