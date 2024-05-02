#importando as bibliotecas
import win32com.client as win32
from datetime import datetime as dt
import pandas as pd
from time import sleep


#assunto
hoje = dt.now()
if hoje.month-1 == 0: mês =12
else: mes = hoje.month-1
if hoje.month-1 ==0: ano = hoje.year-1
else: ano = hoje.year
assunto_data = mes, ano
serviço = input("insira 1 para Escrita Fiscal ou 2 para Departamento de Pessoal")
if serviço == "1": assunto= "A81 - Novidades Urgentes", assunto_data
else: assunto= "Departamento de Pessoal", assunto_data





#importando a tabela
import pandas as pd
tabela_5 = pd.read_csv("clientes_vf2.csv", sep=";")



#criar um email
for y in range(1, 200):
    for x in tabela_5.index:
        if tabela_5.loc[x, "COD"] == y:
            destinatário = (tabela_5.loc[x, "EMAIL"])
            cliente = (tabela_5.loc[x, "CLIENTE"])
            servico = (tabela_5.loc[x, "SERVICO"])
            arquivo = (tabela_5.loc[x, "ARQUIVO"])
            corpo = ("Olá, ", cliente," aqui é o Yuri Becaleti, seu contador. Em busca de melhorar ainda mais",
            "o atendimento que lhe presto desenvolvi um código de computador que envia automaticamente os emails"
            "que antes eu te enviava manualmente com as informações do Departamento de Pessoal e Escrita Fiscal.",
            "Dessa forma além de agilizar a parte operacional (e ter mais tempo para te atender) eu consigo minimizar",
            "o risco de enviar emails errados (como muito raro já aconteceu)",
            "Neste primeiro momento estou enviando os links das pastas de serviço e arquivo para todos os clientes.",
            "Por favor, verifique se os links seguintes correspondem às informações da sua empresa. Caso tenha notado",
            "algo estranho, me avise imediatamente. Obrigado e um abraço!",
            "Pasta Serviços: ", servico,
            "Pasta Arquivo: ", arquivo)

            # criando uma integração com pyton e outolool
            outlook = win32.Dispatch('outlook.application')

            email = outlook.CreateItem(0)
            email.To = destinatário
            email.Subject = f" {assunto}"
            email.Body = f" {corpo}"
            email.Send()

            sleep(5)
