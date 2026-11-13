# Roadmap

Backlog ordenado por impacto × esforço × risco. Itens com ✋ precisam de uma decisão do dono do projecto.

| # | Item | Impacto | Esforço | Risco | Notas |
|---|---|---|---|---|---|
| 1 | ✋ **Substituir `pickle` por formato seguro** (JSON com tipos explícitos) | Alto | Médio | Médio | Carregar um `.pkl` adulterado executa código. A troca exige migração: ler o `.pkl` antigo uma vez e regravar. Decidir se se aceita manter a leitura de `pickle` só para migrar. |
| 2 | **Separar lógica de interface** em módulos (`dados`, `persistencia`, `exportacao`, `graficos`, `ui`) | Alto | Alto | Médio | Hoje cada função mistura `input`, `print` e cálculo. Os 53 testes protegem a refactorização. Manter `python TermiFolha.py` como entrada. |
| 3 | Importar Excel: inteiros com células vazias ficam "decimal" | Médio | Baixo | Baixo | Limitação do pandas. Opção: perguntar ao utilizador o tipo de cada coluna ambígua. |
| 4 | Confirmar antes de importar (Excel/CSV) se já existir folha | Médio | Baixo | Baixo | Hoje é recuperável com `[U]`. |
| 5 | CSV com vírgula decimal e separador `;` (Excel em português) | Médio | Baixo | Baixo | Hoje `1,5` entra como texto. |
| 6 | Describe e Únicos para todas as colunas de uma vez | Baixo | Baixo | Baixo | O manual promete. |
| 7 | ✋ **Actualizar o manual PDF** | Médio | Baixo | Baixo | A fonte do PDF não está no repositório, por isso não o alterei. |
| 8 | PDF com fonte Unicode (TTF) | Baixo | Médio | Baixo | Obriga a distribuir ou localizar uma fonte. |
| 9 | ✋ **Identificadores em inglês** | Baixo | Alto | Médio | Foi pedido, mas o projecto inteiro usa português e renomear 80 funções geraria um diff enorme. Recomendo fazê-lo em conjunto com o item 2. |
| 10 | Registo (`logging`) de erros de gravação e importação | Baixo | Baixo | Baixo | Hoje só há mensagens no ecrã. |

## O que não vale a pena (ainda)

API, autenticação, webhooks, Docker: o produto é uma ferramenta local de terminal; não há utilizadores concorrentes nem serviço a expor.
