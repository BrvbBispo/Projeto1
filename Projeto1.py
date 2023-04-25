# importar a biblioteca
import pandas
import smtplib
import email.message


# importação da base
tab_controle = pandas.read_excel('Vendas.xlsx')

# faturamento a cada loja
fat_tab = tab_controle[['Loja', 'Valor Final']].groupby('Loja').sum()
print(fat_tab)
print('_' * 50)
# quantidade de lojas
quant_tab = tab_controle[['Loja', 'Quantidade']].groupby('Loja').sum()
print(quant_tab)
print('_' * 50)
#ticket médio
ticket_tab = (fat_tab['Valor Final'] / quant_tab['Quantidade']).to_frame('Ticket Médio')
print(ticket_tab)
print('_' * 50)

# enviar email

def enviar_email():

    corpo_email ="""
    <p>Prezado,</p>
    <p>Segue em anexo o relatório:</p>
    <p>Faturamento</p>
    <p>{}</p>
    <p>Quantidade de Lojas</p>
    <p>{}</p>
    <p>Ticket Médio</p>
    <p>{}</p>""".format(fat_tab.to_html(), quant_tab.to_html(), ticket_tab.to_html())


    msg = email.message.Message()
    msg['Subject'] = 'Subject'# assunto do email
    msg['From'] = 'From'# email da pessoa que vai receber
    msg['To'] = 'To'# email da pessoa que vai enviar
    password = 'Password' # senha do email, não é a senha normal, se usa a senha criada para app's

    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado!')

enviar_email()