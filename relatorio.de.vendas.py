import pandas as pd

tabela_vendas = pd.read_excel('vendas.xlsx')
print(tabela_vendas)

#faturamento de cada loja
faturamento = tabela_vendas[['ID Loja', 'Valor final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade de produtos vendidos em cada loja
num_vendas = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(num_vendas)

#ticket médio por produto de cada loja
ticket_medio = faturamento['Valor final'] / num_vendas['Quantidade'].to_frame()
print(ticket_medio)

#enviar um email com o relatório de vendas