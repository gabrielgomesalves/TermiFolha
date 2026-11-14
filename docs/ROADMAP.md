# Roadmap

Estado do backlog definido na auditoria. A aplicação mantém-se num **único ficheiro** por decisão do dono do projecto.

## Concluído

| Item | Resolução |
|---|---|
| `pickle` como formato de estado | Estado em JSON (escrita atómica). O `.pkl` antigo é migrado com um `Unpickler` restrito que recusa `os.system`, `pandas.eval`, etc. |
| Importar sem confirmar | Excel e CSV pedem confirmação se já existir folha. |
| Excel: inteiros com células vazias ficavam decimais | O programa pergunta (só quando há ambiguidade). |
| CSV com vírgula decimal e `;` | Importação reconhece-os; a exportação tem formato Excel PT. |
| Describe e Únicos para todas as colunas | Enter = todas. |
| PDF com fonte Unicode | DejaVu Sans (vem com o matplotlib) com `fpdf2`; Latin-1 com o `fpdf` antigo. |
| Identificadores em inglês | Ver `RENOMEACOES.md`. |
| Registo de erros | `.termifolha.log` rotativo, configurado só em `main()`. |
| CI no GitHub | `.github/workflows/ci.yml` corre o lint e os testes a cada alteração; passou na primeira execução (Novembro de 2026). |
| Testes em Python 3.13 e 3.15 | Passam com os pacotes mais recentes (pandas 3.1, numpy 2.5, matplotlib 3.11, fpdf2 2.9). |
| Manual desactualizado | `docs/MANUAL.md`; o PDF é regenerado dele (`docs/gerar_manual_pdf.py`, desde a v1.3). |

## Concluído na v1.4

Referências relativas (`anterior`, `seguinte`, `desloca`) e entre folhas nas fórmulas, com recálculo de todo o livro por dependências; título e rótulos dos gráficos guardados com a folha; importação de URLs e de bases de dados (SQLite e, com o SQLAlchemy, outras) no menu e na linha de comandos.

## Concluído na v1.3

Fórmulas com referências a células e intervalos (`B2`, `B2:B9`) e agregados (`soma`, `media`, `contar`, `min`, `max`), título e rótulos de gráficos configuráveis e manual em PDF regenerado a partir de `MANUAL.md`.

## Concluído na v1.2

Colunas calculadas, procurar/substituir, converter tipo, várias folhas, juntar/empilhar, tabela dinâmica, desfazer configurável, ver por páginas, colunas visíveis, regras de validação, formatação de valores, correlação, outliers, normalizar, faixas, gráficos em PNG, JSON, colar dados, Markdown/HTML, projectos e modo de linha de comandos. Ver `CHANGELOG.md`.

## Em aberto

| Item | Impacto | Esforço | Notas |
|---|---|---|---|
| Importar de URLs com autenticação (cabeçalhos, tokens) | Baixo | Médio | Hoje só há `GET` simples; para dados protegidos descarrega-se o ficheiro. |
| Referências relativas a células com o estilo `B[-1]` | Baixo | Médio | Hoje faz-se com `anterior(x)`/`desloca(x, n)`. |
| Nomes das funções de teste em inglês | Baixo | Médio | Hoje descrevem o comportamento em português. |
| Fonte Unicode no `fpdf` antigo | Baixo | Médio | Evitado de propósito: o `fpdf` 1.7 escreve ficheiros de cache na pasta da fonte. Basta usar `fpdf2`. |
| Separar `TermiFolha.py` em módulos | Médio | Alto | **Fora do âmbito por decisão do dono**: a aplicação fica num único ficheiro. Os testes já a protegeriam se esta decisão mudar. |

## O que não vale a pena (ainda)

API, autenticação, webhooks, Docker: o produto é uma ferramenta local de terminal.
