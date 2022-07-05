import pandas as pd
import smtplib
import ssl

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
port = 587

smtp_remetente = "smtp-mail.outlook.com"

remetente = "teste.codigopy@outlook.com"

destinatario = "loja.amazing.things@gmail.com"

senha_remetente = "codigopy.1"

mensagem = """

Subject: mensagem teste

"""

SSL_context = ssl.create_default_context()

with smtplib.SMTP(smtp_remetente, port) as server:

    server.starttls(context=SSL_context)

    server.login(remetente,senha_remetente)

    server.sendmail(remetente, destinatario, mensagem)
