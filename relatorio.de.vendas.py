#para ler o arquivo do excel
import pandas as pd
#para conseguir peronalizar o email e enviar pelo outlook
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
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
ticket_medio = (faturamento['Valor final'] / num_vendas['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})   #para colocar o nome 'Ticket Medio' na coluna da tabela
print(ticket_medio)

#enviar um email com o relatório de vendas
port = 587 #é a porta para usar o outlook
server = "smtp-mail.outlook.com"   #link do servidor (sempre será esse)
De = "coloque-seu-email@outlook.com"
Para = "coloque-o-email-de-quem-vai-receber@gmail.com"
senha = "coloque a senha do seu email"
msg = MIMEMultipart()

#corpo do email em html
html_message = f"""
<p>Prezado(a),</p>
<p>Segue abaixo o relatório de vendas de cada loja.</p>
<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{num_vendas.to_html()}

<p>Ticket Médio:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Atte.,</p>
<p>seu nome</p>
"""

msg['From']= De
msg['To']= Para
msg['Subject']="Relatório de Vendas"
msg.attach(MIMEText(html_message, "html"))
text = msg.as_string()
SSLcontext = ssl.create_default_context()

with smtplib.SMTP(server, port) as server:

    server.starttls(context=SSLcontext)

    server.login(De, senha)

    server.sendmail(De, Para, text)