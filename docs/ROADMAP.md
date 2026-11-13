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
| Manual desactualizado | `docs/MANUAL.md`. O PDF não foi alterado: a fonte não está no repositório. |

## Em aberto

| Item | Impacto | Esforço | Notas |
|---|---|---|---|
| Regenerar o manual em PDF a partir de `MANUAL.md` | Baixo | Baixo | Opcional. |
| Nomes das funções de teste em inglês | Baixo | Médio | Hoje descrevem o comportamento em português. |
| Fonte Unicode no `fpdf` antigo | Baixo | Médio | Evitado de propósito: o `fpdf` 1.7 escreve ficheiros de cache na pasta da fonte. Basta usar `fpdf2`. |
| Separar `TermiFolha.py` em módulos | Médio | Alto | **Fora do âmbito por decisão do dono**: a aplicação fica num único ficheiro. Os testes já a protegeriam se esta decisão mudar. |
| CI nunca executada | Médio | Baixo | `.github/workflows/ci.yml` não correu num repositório real. |
| Testar noutras versões de Python | Médio | Baixo | Só foi testado em 3.13. |

## O que não vale a pena (ainda)

API, autenticação, webhooks, Docker: o produto é uma ferramenta local de terminal.
