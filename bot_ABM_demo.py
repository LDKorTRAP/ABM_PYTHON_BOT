import smtplib
import email.message
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from time import sleep
from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

#Contador para o While
cont = int(0)

#função para enviar email senha: qvsh wxvs mrmi whin 
def enviar_email():

    corpo_email = """
    <p>Segue em anexo arquivos de tabela Excel</p>
    <p>- Dados ABM em ordem - 8h - 14h - 18h</p>
    """

    msg = MIMEMultipart()
    msg['Subject'] = "Dados Rastreio ABM"
    msg['From'] = '' #EMAIL REMETENTE
    msg['To'] = '' #EMAIL DESTINATÁRIO
    password = '' #CHAVE DE APP GERADA PELO GMAIL
    msg.attach(MIMEText(corpo_email, 'html'))


    filename = 'M3Risk.xlsx'

    attachment = open('CAMINHO DO LOCAL DOS ARQUIVOS PÓS DOWNLOAD','rb') #EX: C:\\Users\\ldkse\\Downloads\\M3Risk.xlsx

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)

    filename = 'M3Risk(1).xlsx'
    attachment = open('CAMINHO DO LOCAL DOS ARQUIVOS PÓS DOWNLOAD','rb') #EX: C:\\Users\\ldkse\\Downloads\\M3Risk(1).xlsx

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)

    filename = 'M3Risk(2).xlsx'

    attachment = open('CAMINHO DO LOCAL DOS ARQUIVOS PÓS DOWNLOAD','rb') #EX: C:\\Users\\ldkse\\Downloads\\M3Risk(2).xlsx

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string())
    print('Email enviado')


#iniciando bot via firefox
servico = Service(GeckoDriverManager().install())
navegador = webdriver.Firefox(service = servico)
navegador.maximize_window()

#ACESSANDO A PLATAFORMA ABM
navegador.get("http://abm.m3risk.com.br/login")
sleep(10)

#ACESSANDO PAINEL LOGÍSTICO E BAIXANDO PLANILHA
navegador.find_element(By.XPATH, '/html/body/div[2]/div[4]/form/div[1]/input').send_keys("") #EMAIL DE LOGIN AQUI
navegador.find_element(By.XPATH, '/html/body/div[2]/div[4]/form/div[2]/input').send_keys("") #SENHA AQUI
navegador.find_element(By.XPATH, '/html/body/div[2]/div[4]/form/div[3]/div[2]/button').click()
sleep(5)
navegador.find_element(By.XPATH, '/html/body/div[2]/aside/section/ul/li[5]/a').click()
sleep(10)

#horarios: 8, 14, 18
while True:
    navegador.refresh()
    tempo = int('%d' % (datetime.now().hour))
#condições para executar o download do arquivo excel em 3 horários dados
    if cont < 3:
        if tempo == 8:
            if cont == 0:
                navegador.refresh()
                sleep(10)
                navegador.find_element(By.XPATH, '/html/body/div[2]/div/section[2]/div[1]/div/div/div/div/div[1]/div[1]/a[2]').click()
                sleep(10)
                cont += 1 

        elif tempo == 14:
            if cont == 1:
                navegador.refresh()
                sleep(10)
                navegador.find_element(By.XPATH, '/html/body/div[2]/div/section[2]/div[1]/div/div/div/div/div[1]/div[1]/a[2]').click()
                sleep(10)               
                cont += 1

        elif tempo == 18:
            if cont == 2:
                navegador.refresh()
                sleep(10)
                navegador.find_element(By.XPATH, '/html/body/div[2]/div/section[2]/div[1]/div/div/div/div/div[1]/div[1]/a[2]').click()
                sleep(10)
                cont += 1

    else: 
        #enviando via email
        enviar_email()
        break
