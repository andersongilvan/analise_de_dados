import pandas as pd
import win32com.client as win32


def MostarLinha():
    print('-' * 50)





tabela_vendas = pd.read_excel('vendas.xlsx')
MostarLinha()
# visualizar a base de dados
pd.set_option('display.max_columns', None)

#faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(f'\n\nfaturamento por loja\n')
print(faturamento)
MostarLinha()

# quantidade de produto vendido por loja
print(f'\n\nquantidade de produto vendido por loja\n')
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
MostarLinha()

# Calcular a média de preço de cada produtos
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(f'\n\nCalculando a média de preço de cada produtos\n')

ticket_medio = ticket_medio.rename(columns={0: 'ticket medio'})

print(ticket_medio)



# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'torresdinho8@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f''' 
<p> Presados,</p>

<p> Segue o relatório de vendas por cada loja. </p>

<p> Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>  Quantidade Vendida </p>
{quantidade.to_html()}

<p> Tiket Médio dos produtos por cada Loja </p>
{ticket_medio.to_html(formatters={'ticket Médio':'R${:,.2f}'.format})}


<p> Qualquer dúvida estou à disposição. </p>

<p> Att,, <p/>
<p> Anderson </p>
'''

mail.Send()

print('\n\nEmail enviado')












