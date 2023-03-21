# Automação-de-email
## Objetivo

Este código tem como objetivo ler e tratar dados de uma base de dados, e enviar automaticamente um email com um relatório sobre o arquivo da base de dados.

## Como funciona

Utilizando a IDE Pycharm, o código na linguagem python recebe uma base de dados do arquivo “Vendas” na qual possuiu os dados das vendas de produtos de várias lojas como:

- Data
- ID da loja
- Produto
- Quantidade
- Valor unitário
- Valor Final

Após a leitura dos dados, o código calcula o faturamento por loja, a quantidade de produtos vendidos por loja e o ticket médio por produto em cada loja.

Por fim, o código envia um email automaticamente com o relatório possuindo os dados tratados em tabelas para o endereço desejado.

## Bibliotecas usadas

Foram usadas as bibliotecas **Pandas**, **Openpyxl**, e **Win32com.client.**

## Detalhes

- É necessário deixar o arquivo “Vendas” na mesma pasta do código para que ele seja lido
- É necessário ter o Outlook baixado no PC para poder enviar o email
- Na linha de código abaixo onde está “seuemailaqui@gmail.com” você substitui pelo email na qual você deseja enviar o relatório

```python
mail.To = 'seuemailaqui@gmail.com' #Endereço de email que vai enviar
```
