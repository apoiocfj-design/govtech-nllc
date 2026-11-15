# Catálogo Nacional de Bens e Serviços (CNBS)

Este repositório permite consultar os itens dos catálogos de materiais e serviços utilizados pelo ComprasGov.

## Arquivos disponíveis

- `catmat.csv`: itens do Catálogo de Materiais (CATMAT).
- `catser.csv`: itens do Catálogo de Serviços (CATSER).

Os dois arquivos incluem itens ativos e inativos. A situação é informada na coluna `statusItem` (`catmat.csv`) e na coluna `statusServico` (`catser.csv`): `True` indica item ativo e `False`, item inativo.

## Organização e atualização

O catálogo não possui particionamento por data: os dois arquivos permanecem diretamente na raiz deste repositório. A atualização ocorre semanalmente, aos sábados, às 6h.

Os arquivos são substituídos integralmente a cada atualização. Dessa forma, itens podem ser incluídos ou ter suas informações alteradas após uma publicação anterior. Para consultas atuais, utilize sempre a versão mais recente disponível.

## Como utilizar

Baixe diretamente o CSV desejado. Os arquivos usam codificação UTF-8, possuem os nomes das colunas na primeira linha e utilizam ponto e vírgula (`;`) como separador. Não é necessário descompactá-los.
