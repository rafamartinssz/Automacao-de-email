import pandas as pd
import openpyxl
import win32com.client as win32

#Importar a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")

#Visualizar a base de dados
pd.set_option('display.max_columns',None)    #Define como serão exibidas as colunas
print(tabela_vendas)

#Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()  #Calcula o faturamento total por loja
print(faturamento)

#Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()    #Calcula a quantidade total de produtos vedidos por loja
print(quantidade)

print('-' * 50)
#Ticket médio por produto em cada loja

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()  #Calcula o ticket médio por produto em cada loja
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})                    #Renomeia a coluna com o ticket médio calculado
print(ticket_medio)

#Enviar email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'rafamartinssz18@gmail.com' #Endereço de email que vai enviar
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Segue abaixo o relatório de vendas por loja</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendido:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

'''
mail.Send()
