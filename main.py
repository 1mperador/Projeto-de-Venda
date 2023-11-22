import pandas as pd
import win32com.client as win32

# Importar a base de dados 
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Vizualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('-' * 50)
# Faturamento por Loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' * 50)
# Quantidade de produtos vendidos por loja 
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# Ticket médio por produto em cada loja 
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# Enviar um email com o relatorio
outlook = win32.Dispatch('outlook.aplication')
mail = outlook.CreateItem(0)
mail.To = 'email'
mail.Subject = 'Relatorio de vendas por loja '
mail.HTMLBody = '''
Prezados,

Segue o Relatorio de vendas por cada loja.

Faturamento:
{}

Quantidade Vendida:
{}

Ticket Media dos Produtos em cada Loja:
{}

Qualquer dúvida estou á disposição.

Att.,
_user_
'''


mail.Send()
