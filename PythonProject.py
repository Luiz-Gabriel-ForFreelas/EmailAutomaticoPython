# O arquivo não contém credenciais reias a fim de manter a segurança da informação utilizada no projeto efetivo.

# importação de bibliotecas
import pandas as pd
import psycopg2
from datetime import date
from calendar import monthrange
from openpyxl import Workbook, load_workbook
import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# PARTE 1 LIMPAR A PLANILHA ANTIGA  ==================================================
# acessa o excel desejado
excel = load_workbook("exemplo.xlsx")

# define a aba ativa do excel
aba_ativa = excel.active

# colunas
colunas = ["A", "B", "C", "D", "E", "F", "G", "H"]

for coluna in colunas:
    for num, celula in enumerate(aba_ativa[coluna]):
        aba_ativa[f"{coluna}{num + 2}"] = ""

# PARTE 2 RALIZAR O SELECT NO POSTGRES  ==================================================
# descobre os dias que o mês tem
data_atual = date.today()

# monthrange retorna o último dia do mês
ultimo_dia = f"{date.today().year}-{date.today().month - 1}-{monthrange(data_atual.year, data_atual.month - 1)[1]} 23:59:59"

# indica o dia 1 do mês passado
primeiro_dia = f"{date.today().year}-{date.today().month - 1}-1 00:00:00"

# define a query a ser executada
postgreSQL_select_Query = f""" SELECT montado para obter as informações entre o dia {primeiro_dia} e {ultimo_dia} """

# estabelece conexão com a base de dados
conexao = psycopg2.connect(database='base_selecionada',
                           host='endereço',
                           user='usuario',
                           password='senha',
                           port='porta')

# extende a conexão para receber comandos de query
cursor = conexao.cursor()

# executa a query dentro da base de dados
cursor.execute(postgreSQL_select_Query)

# armazena os resultados da query
mobile_records = cursor.fetchall()

# roda pelas linhas da tabela e as redefine com os resultados da query por linha
for num, row in enumerate(mobile_records):
    aba_ativa[f"A{num + 2}"] = row[0]
    aba_ativa[f"B{num + 2}"] = row[1]
    aba_ativa[f"C{num + 2}"] = row[2]
    aba_ativa[f"D{num + 2}"] = row[3]
    aba_ativa[f"E{num + 2}"] = row[4]
    aba_ativa[f"F{num + 2}"] = row[5]
    aba_ativa[f"G{num + 2}"] = row[6]
    aba_ativa[f"H{num + 2}"] = int(row[7]) # força a definição do item como inteiro para garantir uma boa visualização final

# encerra a conecão com o servidor
cursor.close()
conexao.close()

# fazer o refresh
ws = excel["Sheet"] # seleciona o sheet resumo
pivot = ws._pivots[0] # define o cache a ser atualizado
pivot.cache.refreshOnLoad = True # marca como verdadeiro a atualização automática do cache ao ser aberto

# salva o arquivo em um xlsx separado, para manter o arquivo base estático
excel.save("exemplo_final.xlsx")

# PARTE 3 ENCAMINHAR O EMAIL PARA O RH ==================================================

# Monagem do email

# variaveis de envio
login = "seueamil@gmail.com.br"
senha = "senha do email"

# define o corpo, cabeçalho, destinatário e tipo de leitura do email.
corpo = f"Mensagem que será recebida pelo destinatário."
email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = "destinatario@gmail.com.br"
email_msg['Subject'] = "Cabeçalho do email"
email_msg.attach(MIMEText(corpo, 'html'))

# Anexando arquivo ao email "exemplo_final.xlsx"
# encontra o arquivo e o lê como binário.
arquivo = "exemplo_final.xlsx"
attachment = open(arquivo, 'rb')

# verifica o arquivo com a base64 para o encaminhamento em bits
att = MIMEBase('application', 'octet-stream')
att.set_payload(attachment.read())
encoders.encode_base64(att)

# define o header do anexo.
att.add_header('Content-Disposition', 'attachment; filename=boletos.xlsx')

# encerra a leitura do anexo.
attachment.close()

# anexa o arquivo ao email
email_msg.attach(att)

# Conecatando ao servidor de emails
host = "host de emails"
port = "porta do host"
server = smtplib.SMTP(host, port)
server.ehlo()
server.starttls()
server.login(login, senha)

# Enviando o email e encerrando a conexão
server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
server.quit()
