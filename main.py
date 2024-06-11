import win32com.client as win32

# importar a BASE DE DADOS
import openpyxl as openpyxl
import pandas as pd

# organizar a BASE DE DADOS
tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos VENDIDOS por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# ticket médio (faturamento / produtos por loja)
# preço médio do produto de cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# # enviar um email com o RELATÓRIO
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
# Escreva o destinatário aqui:
mail.To = 'email@example.com'
mail.subject = 'Relatório de vendas por loja.'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Preço médio do produto de cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>At.te,</p>
<p>Iago</p>
'''


mail.Send()
print('Email enviado!')