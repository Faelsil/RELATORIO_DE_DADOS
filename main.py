# -*- coding: utf-8 -*-
"""
Created on Tue Jul 30 10:36:15 2024

@author: Rafae
"""
#importar a base de dados - OK
#visualizar a base de dados - OK
#faturamento por loja - OK
#quantidade de produtos vendidos por loja - OK
#ticket médio por produto em cada loja - OK
#enviar um email com o relatório - OK

import pandas as pd
import smtplib
import email.message

#importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualizar a base de dados
pd.set_option('display.max_columns', None)
#print(tabela_vendas)

#faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#print(faturamento)

#quantidade de produtos vendidos por loja
qtdProduto = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#print(qtdProduto)

#ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtdProduto['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
#print(ticket_medio)

{}
def enviar_email():  
    corpo_email = f'''
    <h1><b>Relatorio de vendas</b></h1>
    <h3><b>Segue o relatorio de vendas por cada loja</b></h3>
    
    <p><b>Faturamento</b>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}</p>
    
    <p><b>Quantidade</b>
    {qtdProduto.to_html()}</p>
    
    <p><b>Ticket Médio dos produto em cada loja</b>
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}</p>
    
    <p>Att.,</p>
    <p>Rafael Aparecido</p>
    '''

    msg = email.message.Message()
    msg['Subject'] = "Relatorio de vendas por loja"
    msg['From'] = 'rafaelsilva.ap94@gmail.com'
    msg['To'] = 'rafaelsilva.ap94@gmail.com'
    password = 'p q x n q q n o g j v s x z e o' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()
