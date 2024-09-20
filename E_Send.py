import smtplib
import email.message
from datetime import datetime,date
import os
import pandas as pd


diretorio_principal = os.path.dirname(__file__)
diretorio_arquivos = os.path.join(diretorio_principal, 'arquivos')
relatorio = os.path.join(diretorio_arquivos, 'relatorio_erro.txt')
diretorio_cad = os.path.join(diretorio_arquivos, 'cadastros')
cadastros = os.path.join(diretorio_cad,"cadastros.xlsx")

dt = datetime.now()
hora = dt.hour
minuto = dt.minute
segundo = dt.second
dia = dt.day
mes = dt.month
ano = dt.year
hora_format = hora
minuto_format = minuto
segundo_format = segundo

mail_key = "iuybulnjetvhnvdk"
mail = ''
dev_mail = ''

def Obter_Hora():

    dt = datetime.now()
    hora = dt.hour
    minuto = dt.minute
    segundo = dt.second
    dia = dt.day
    mes = dt.month
    ano = dt.year
    hora_format = hora
    minuto_format = minuto
    segundo_format = segundo
    
    if hora <10:
        hora_format = f'0{hora}'
    if minuto <10:
        minuto_format = f'0{minuto}'
    if segundo <10:
        segundo_format = f'0{segundo}'

    h = f'{hora_format}:{minuto_format}'

    return h

def Dia():
    dt = datetime.now()
    d = date(dt.year,dt.month,dt.day)
    hora = dt.hour
    minuto = dt.minute
    segundo = dt.second
    dia = dt.day
    mes = dt.month
    ano = dt.year
    dia_format = dia + 1
    mes_format = mes
    ano_format = ano
    semana = d.isoweekday()

    if(mes == (1,3,5,7,8,10,12)):
        if(dia == 31):
            dia_format = 1
        else:
            dia_format = dia + 1
    if(mes == (4,6,9,11)):
        if(dia == 30):
            dia_format = 1
        else:
            dia_format = dia + 1
    if(mes == 2):
        if(dia == 28 or dia == 29):
            dia_format = 1
        else:
            dia_format = dia + 1
                
    formatado = f'{dia_format}/{mes}/{ano}'
    return formatado

def Enviar_Email(aluno,destino):
    try:
        print(f"Enviando E-Mail para {aluno} | {destino}\n")
        remetente = mail
        msg = email.message.EmailMessage()
        msg['Subject']  = 'Sirius ALBOT v1.5 | Formulário do Almoço *Destino*'
        msg['From'] = remetente
        msg['To'] = destino
        corpo_email = f"""
            <div class='container' style='justify-content: center;
    align-items: start;
    
    width: 100%;
    height: 300px ;
    border-radius: 50px;'>

        <h3 style='display: flex; font-family: Arial; margin: 10px;'>Sirius ALBOT v1.5 [{Obter_Hora()}]</h3>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Caro(a) {aluno}. Cadastro no Formulário de almoço para o dia [{Dia()}] concluído.</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Dúvidas? siriuscontactmail@gmail.com</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Atenciosamente Equipe de Desenvolvimento Sirius.</span>

    </div>
        """
        
        corpo_email = corpo_email.encode('utf-8')
        aluno = aluno.encode('utf-8')
        destino = destino.encode('utf-8')


        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            
            smtp.login(mail, mail_key)
            smtp.send_message(msg)
    except Exception as Ex:
        print(f"Houve um erro ao tentar enviar o Email...")
        
        with open(relatorio, 'a') as arquivo:
            arquivo.write(f'[{dia}/{mes}/{ano} {hora}:{minuto}]\n')
            arquivo.write(f"{Ex}")

def Enviar_Email_Erro(aluno,destino):
    try:
        print(f"Enviando E-Mail para {aluno} | {destino}\n")
        remetente = mail
        msg = email.message.EmailMessage()
        msg['Subject']  = 'Sirius ALBOT v1.5 | Formulário do Almoço *Destino*'
        msg['From'] = remetente
        msg['To'] = destino
        corpo_email = f"""
        <div class='container' style='justify-content: center;
        align-items: start;
        width: 100%;
        height: 300px ;
        border-radius: 50px;'>

        <h3 style='display: flex; font-family: Arial; margin: 10px;'>Sirius ALBOT v1.5 [{Obter_Hora()}]</h3>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Caro(a) {aluno}. Ocorreu um erro ao preencher sua solicitação de almoço do dia [{Dia()}].</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Pedimos desculpas pelo inconveniente e gostaríamos de relatar que nossa equipe já está trabalhando para a solução do problema.</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Dúvidas? siriuscontactmail@gmail.com</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Atenciosamente Equipe de Desenvolvimento Sirius.</span>

    </div>
        """
        
        
        corpo_email = corpo_email.encode('utf-8')
        aluno = aluno.encode('utf-8')
        destino = destino.encode('utf-8')


        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            
            smtp.login(mail, mail_key)
            smtp.send_message(msg)
            
    except Exception as Ex:
        print(f"Houve um erro ao tentar enviar o Email...")
        
        with open(relatorio, 'a') as arquivo:
            arquivo.write(f'[{dia}/{mes}/{ano} {hora}:{minuto}]\n')
            arquivo.write(f"{Ex}")

def Relatorio_Erro(aluno,destino,erro):
    try:
        print(f"Enviando Relatório de Erro.\n")
        remetente = mail
        msg = email.message.EmailMessage()
        msg['Subject']  = 'Sirius ALBOT v1.5 | Relatório de Erro'
        msg['From'] = remetente
        msg['To'] = dev_mail
        corpo_email = f"""
        <div class='container' style='justify-content: center;
        align-items: start;
        width: 100%;
        height: 300px ;
        border-radius: 50px;'>

        <h3 style='display: flex; font-family: Arial; margin: 10px;'>Sirius ALBOT v1.5 [{Obter_Hora()}]</h3>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Na tentativa de fazer cadastro do aluno {aluno} | {destino}, foi encontrado um erro e o sistema entrou em Except. Segue abaixo o relatório de erro</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>{erro}</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Atenciosamente Equipe de Desenvolvimento Sirius.</span>

    </div>
        """
        
        corpo_email = corpo_email.encode('utf-8')
        aluno = aluno.encode('utf-8')
        destino = destino.encode('utf-8')


        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            
            smtp.login(mail, mail_key)
            smtp.send_message(msg)
    except Exception as Ex:
        print(f"Houve um erro ao tentar enviar o Email...")
        
        with open(relatorio, 'a') as arquivo:
            arquivo.write(f'[{dia}/{mes}/{ano} {hora}:{minuto}]\n')
            arquivo.write(f"{Ex}")

def Relatorio_Conclusao(quantidade):
    try:
        print(f"Enviando Relatório de Conclusão.\n")
        remetente = mail
        msg = email.message.EmailMessage()
        msg['Subject']  = 'Sirius ALBOT v1.5 | Relatório de Conclusão Formulário de Almoço.'
        msg['From'] = remetente
        msg['To'] = dev_mail
        corpo_email = f"""
        <div class='container' style='justify-content: center;
        align-items: start;
        width: 100%;
        height: 300px ;
        border-radius: 50px;'>

        <h3 style='display: flex; font-family: Arial; margin: 10px;'>Sirius ALBOT v1.5 [{Obter_Hora()}]</h3>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Relatório de conclusão. Sirius ALBOT completou com sucesso o cadastro de {quantidade} alunos do dia [{Dia()}].</span>
        <span style='display: flex; font-family: Arial; margin: 10px;'>Atenciosamente Equipe de Desenvolvimento Sirius.</span>

    </div>
        """
        
        corpo_email = corpo_email.encode('utf-8')
        


        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email)
        with open(cadastros, 'rb') as cadastros_a:
            cad = cadastros_a.read()
            msg.add_attachment(cad, maintype='application', subtype='xlsx',filename="cadastros.xlsx")


        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            
            smtp.login(mail, mail_key)
            smtp.send_message(msg)
    except Exception as Ex:
        print(f"Houve um erro ao tentar enviar o Email...")
        
        with open(relatorio, 'a') as arquivo:
            arquivo.write(f'[{dia}/{mes}/{ano} {hora}:{minuto}]\n')
            arquivo.write(f"{Ex}")

def Enviar_Cadastro(aluno,destino):
    try:
        print(f"Enviando E-Mail para {aluno} | {destino}\n")
        remetente = mail
        msg = email.message.EmailMessage()
        msg['Subject']  = 'Sirius ALBOT v1.5 | Cadastro.'
        msg['From'] = remetente
        msg['To'] = destino
        corpo_email = f"""
            <p>
            <h3>
            <b>Sirius ALBOT v1.5 [{Obter_Hora()}]</b>
            </h3>
            </p>
            <p>Caro(a) {aluno}. Você foi cadastrado como usuário no Sirius ALBOT v1.5.</p>
            <p>A partir de hoje, suas solicitações de almoço em: *Destino*, estarão sendo preenchidas pelo mesmo.</p>
            <p>Solicitamos que, caso for faltar, envie um aviso até no máximo às 23h do dia pré anterior a solicitação.</p>
            <p>Exemplo: Meu cadastro é para a Quinta Feira, logo, caso eu venha a faltar, devo tentar avisar até às 23h da terça.</p>
            <p>Caso ocorra de você adoecer e seu almoço ainda for solicitado, solicite o cancelamento direto com a organização do restaurante.</p>
            <p>Este é um E-Mail automático emitido pelo sistema, caso você não tenha solicitado o mesmo, entre em contato para cancelamento de cadastro</p>
            <p>Dúvidas? {mail}</p>
            <p>Atenciosamente Equipe de Desenvolvimento Sirius.</p>
        """
        
        corpo_email = corpo_email.encode('utf-8')
        aluno = aluno.encode('utf-8')
        destino = destino.encode('utf-8')


        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            
            smtp.login(mail, mail_key)
            smtp.send_message(msg)
    except Exception as Ex:
        print(f"Houve um erro ao tentar enviar o Email...")
        
        with open(relatorio, 'a') as arquivo:
            arquivo.write(f'[{dia}/{mes}/{ano} {hora}:{minuto}]\n')
            arquivo.write(f"{Ex}")

