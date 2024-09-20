#Gabriel Araújo dos Santos 20/02/2024
#Sirius ALBOT v1.5| siriuscontactmail@gmail.com

import pandas as pd
from time import sleep
from playwright.sync_api import sync_playwright
import os
from datetime import date,datetime
import smtplib
import email.message
import E_Send as mail
import speech_recognition as sr 
import pyttsx3
import pyautogui as pag


#TURMAS
# 1º Ano – G1 : 1
# 1º Ano – G2 : 2
# 2º Ano – G1 : 3
# 2º Ano – G2 : 4
# 3º Ano – G1 e G2 : 5

#CURSOS
# Integrado – Edificações : 1
# Integrado – Geologia : 2 
# Integrado – Informática : 3 
# Integrado – Mineração : 4

#DIAS
# Segunda: 1
# Terça: 2
# Quarta: 3
# Quinta: 4
# Sexta: 5
# Sábado: 6

qt = 0
dt = datetime.now()
d = date(dt.year,dt.month,dt.day)
hora = dt.hour
minuto = dt.minute
segundo = dt.second
dia = dt.day
mes = dt.month
ano = dt.year
dia_format = dia
mes_format = mes
ano_format = ano
semana = d.isoweekday()

url = "https://forms.gle/b48LTTGSLjabEKXn8"



diretorio_principal = os.path.dirname(__file__)
diretorio_arquivos = os.path.join(diretorio_principal, 'arquivos')
diretorio_cad = os.path.join(diretorio_arquivos, 'cadastros')
dados = os.path.join(diretorio_arquivos, 'dados.xlsx')
relatorio = os.path.join(diretorio_arquivos, 'relatorio_erro.txt')
cadastros = os.path.join(diretorio_cad,f"cadastros.xlsx")

app_data_path = os.getenv("LOCALAPPDATA")
user_data = os.path.join(app_data_path, 'Google\\Chrome\\User Data\\Default')


df = pd.read_excel(dados)

class Selectors():
    Email = "#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(1) > div > div > div.AgroKb > div > div.aCsJod.oJeWuf > div > div.Xb9hP > input"
    Nome = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input'
    Matricula = '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(3) > div > div > div.AgroKb > div > div.aCsJod.oJeWuf > div > div.Xb9hP > input'
    Curso_Click = '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(4) > div > div > div.vQES8d > div > div:nth-child(1) > div.ry3kXd > div.MocG8c.HZ3kWc.mhLiyf.LMgvRb.KKjvXb.DEh1R'
    Turma_Click = '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(5) > div > div > div.vQES8d > div > div:nth-child(1) > div.ry3kXd > div.MocG8c.HZ3kWc.mhLiyf.LMgvRb.KKjvXb.DEh1R'
    C_Turno = '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(7) > div > div > div.oyXaNc > div > div > span > div > div:nth-child(1) > label'
    Data_Click = '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(6) > div > div > div:nth-child(2) > div > div > div.rFrNMe.yqQS1.hatWr.zKHdkd > div.aCsJod.oJeWuf > div > div.Xb9hP > input'
    Data_Send = '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(6) > div > div > div:nth-child(2) > div > div > div.rFrNMe.yqQS1.hatWr.zKHdkd > div.aCsJod.oJeWuf > div > div.Xb9hP > input'
    Send = '#mG61Hd > div.RH5hzf.RLS9Fe > div > div.ThHDze > div.DE3NNc.CekdCb > div.lRwqcd > div > span > span'
Sa = Selectors()


def Cadastrar_Aluno(nome,email):
    aluno = {
      "NOME" : nome,
      "EMAIL" : email,
      "DATA" : [f'{dia}/{mes}/{ano}']
    }

    aluno_DF = pd.DataFrame(aluno)
  
    alunos = pd.DataFrame(pd.read_excel(cadastros,engine="openpyxl"))
    alunos = alunos.dropna()
    alunos_append = pd.concat([alunos,aluno_DF],ignore_index=True, join="inner")
    alunos_append.to_excel(cadastros)


print("Sirius: ALBOT v1.5 Running | siriuscontactmail@gmail.com\n")



for index,row in df.iterrows():
    
    if(row['DIA'] == 0):
        pass
    
    if(semana+1 == row['DIA']):
        
        try:
            with sync_playwright() as p:
                print(f"Index:{str(index)}\nFazendo cadastro de: {row["NOME"]} | {str(row["MATRICULA"])}\n") 
                
                
                browser = p.chromium.launch(headless=False)
                page = browser.new_page()
                page.goto(url)
                class Get():
                    Curso_Edif = page.get_by_role('option', name= "Integrado - Edificações")
                    Curso_Geol = page.get_by_role('option', name= "Integrado - Geologia")
                    Curso_Info = page.get_by_role('option', name= "Integrado - Informática")
                    Curso_Mine = page.get_by_role('option', name= "Integrado - Mineração")
                    Turma_1G1 = page.get_by_role('option', name= "1º Ano - G1")
                    Turma_1G2 = page.get_by_role('option', name= "1º Ano - G2")
                    Turma_2G1 = page.get_by_role('option', name= "2º Ano - G1")
                    Turma_2G2 = page.get_by_role('option', name= "2º Ano - G2")
                    Turma_3G = page.get_by_role('option', name= "3º Ano - G1 e G2")
                G = Get()
                def Curso():
                    
                    if row['CURSO'] == 1:
                        G.Curso_Edif.click()
                        
                    if row['CURSO'] == 2:
                        G.Curso_Geol.click()
                    if row['CURSO'] == 3:
                        G.Curso_Info.click()

                    if row['CURSO'] == 4:
                        G.Curso_Mine.click()

                def Turma():
                    if row["TURMA"] == 1:
                        G.Turma_1G1.click()
                    
                    if row["TURMA"] == 2:
                        G.Turma_1G2.click()
                    
                    if row["TURMA"] == 3:
                        G.Turma_2G1.click()
                    
                    if row["TURMA"] == 4:
                        G.Turma_2G2.click()
                    
                    if row["TURMA"] == 5:
                        G.Turma_3G.click()
            
                            
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
                    
                    formatado = f'{dia_format}{mes}{ano}'
                    return formatado
                        
                def Tab(quantidade):
                    for i in range(quantidade):
                        page.keyboard.press("Tab")
                
                def KTab(quantidade):
                    for i in range(quantidade):
                        pag.press("Tab")

                print("Aguardando página...\n")
                while True:
                    if(page.url[-10:] == "closedform"):
                        page.reload()
                        sleep(3)
                        
                    else:
                        print("Página Disponível... \n")
                        sleep(1)
                        break
                        
                
                qt+= 1
                
                page.fill(Sa.Email,row['EMAIL'])
                sleep(0.05)
                page.fill(Sa.Nome,row['NOME'])
                sleep(0.05)
                page.fill(Sa.Matricula,str(row['MATRICULA']))
                sleep(0.05)
                page.click(Sa.Curso_Click)
                sleep(0.1)
                Curso()
                sleep(0.1)
                page.click(Sa.Turma_Click)
                sleep(0.1)
                Turma()
                sleep(0.1)
                page.click(Sa.C_Turno)
                sleep(0.1)
                page.click(Sa.Data_Click)
                sleep(0.1)
                page.keyboard.press("ArrowLeft")
                sleep(0.1)
                page.keyboard.press("ArrowLeft")
                sleep(0.1)
                page.keyboard.type(Dia())
                sleep(0.1)
                page.click(Sa.Send)
                sleep(0.5)
                

                while True:
                    if(page.url[-12:] == "formResponse"):
                        sleep(1)
                        break
                        
                    else:
                        sleep(1)                                                                                          

                mail.Enviar_Email(f"{row['NOME']}",f"{row['EMAIL']}")

                print(f'{row['NOME']} finalizado com sucesso. Passando para o próximo...\n')
                Cadastrar_Aluno(f'{row["NOME"]}',f'{row["EMAIL"]}')
                sleep(1)
            

                browser.close()
        except Exception as Ex:
            print(f"Na tentativa de fazer cadastro do aluno {row['NOME']} | {row['EMAIL']}, foi encontrado um erro e o sistema entrou em Except. Relatório de Erro Salvo.")
            with open(relatorio, 'a') as arquivo:
                arquivo.write(f'[{dia}/{mes}/{ano} {hora}:{minuto}]\n')
                arquivo.write(f"{Ex}")
            mail.Relatorio_Erro(f"{row['NOME']}",f"{row['EMAIL']}",f"{Ex}")
            mail.Enviar_Email_Erro(f"{row['NOME']}",f"{row['EMAIL']}")
try:  
    print("Sim")          
    mail.Relatorio_Conclusao(qt)
except Exception as Ex:
    print("Erro ao enviar relatório de conclusão, ainda é possivel encontrar a planilha nos arquivos do sistema.")
    with open(relatorio, 'a') as arquivo:
        arquivo.write(f'[{dia}/{mes}/{ano} {hora}:{minuto}]\n')
        arquivo.write(f"{Ex}")
print(f'Finalizado com sucesso. | ALBOT v1.5')

