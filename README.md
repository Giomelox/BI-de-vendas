# Geração de Planilha de Rebaixas para Produtos

## Funcionalidades
Leitura de dados da planilha bi_semanal.xlsx.

Filtragem de clientes com nomes específicos.

Formatação de datas no padrão dd/mm/yyyy.

Criação de nova planilha BI Mix e Novo - Frios e Secos.xlsx com 4 abas organizadas por tipo e cliente.

Geração das planilhas de rebaixa REBAIXA FRIOS NOVO.xlsx; REBAIXA SECOS NOVO.xlsx; REBAIXA FRIOS MIX.xlsx; REBAIXA SECOS MIX.xlsx com:

Agrupamento por cliente.

Fórmulas automáticas de investimento e sell out.

Inclusão de produtos com códigos e preços mapeados de uma planilha base.

Estilo visual com cores, negrito, bordas e alinhamento.

Formatação monetária (R$) para colunas específicas.

## Observações
As funções são responsáveis por gerar as planilhas finais de rebaixas.

O mapeamento de códigos e preços depende do campo SKU Description. É importante garantir que os nomes estejam limpos e padronizados.

A execução atual realiza o mapeamento dentro do loop, o que pode ser otimizado usando .merge() ao final da montagem do DataFrame.

## Como Executar
Instale os pacotes necessários:
- pip install pandas openpyxl
- execute o código
  
Ao final, será gerado o arquivo:

BI Mix e Novo - Frios e Secos.xlsx: Dados filtrados por tipo e cliente.

REBAIXA FRIOS NOVO.xlsx: Tabela formatada de rebaixas dos frios para o cliente NOVO.

REBAIXA SECOS NOVO.xlsx: Tabela formatada de rebaixas dos secos para o cliente NOVO.

REBAIXA FRIOS MIX.xlsx: Tabela formatada de rebaixas dos frios para o cliente MIX.

REBAIXA FRIOS SECOS.xlsx: Tabela formatada de rebaixas dos secos para o cliente MIX.
