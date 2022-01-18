## Projeto de data mining de texto em multíplos arquivos PDF e exportando em excel

Este projeto tem o intuito de analisar e minerar palavras-chave de contratos imobilíarios e salvar esses dados em uma planilha de importação num software ERP.

A motivação da criação desse código veio de um problema de negócio onde era necessário um trabalho manual e repetitivo de seleção de dados de contratos de compra de lotes, onde era necessário repetir esse trabalho para mais de 5000 contratos. Certamente não era uma situação produtiva nem aceitável, uma tentativa de migração de banco de dados havia sido feita sem sucesso.

O código funciona a partir da ideia da leitura e tratamento do texto do contrato, refinamento e procura das palavras-chave necessárias, que são armazenadas num dataframe, que posteriormente é convertido para um modelo excel compátivel com o do ERP.

* Um consultor experiente alimenta o sistema em torno de 1 contrato a cada 3 minutos, com mais dois minutos para as condições de pagamento, totalizando cinco minutos por contrato;
* Na prática se exportava 30 contratos por dia de consultoria, que custa em torno dos R$1600;
* Nesse cliente, estimou-se uma necessidade de 5000 contratos exportados;
* Nos testes realizados, conseguimos realizar automaticamente a exportação de 5000 contratos em menos de 15 minutos;
* Gerando uma economia gigantgesca de tempo e dinheiro para a empresa.

Sendo esse código flexível para mudanças, sendo possível realizar a mineração de dados de diferentes documentos e tipos de arquivos. Quando for necessário buscar certas palavras de um texto e exportá-las numa planilha ou csv.

Obs: Os dados do contrato de exemplo são meramente ilustrativos.

Desenvolvido em conjunto com [Jonatas Salles](https://github.com/jonatas-salles)
