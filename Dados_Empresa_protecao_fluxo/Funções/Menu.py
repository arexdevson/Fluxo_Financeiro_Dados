import datetime
from time import sleep
import pandas as pd
import matplotlib.pyplot as plt
from datetime import *

# envio do email
import win32com.client as win32

# carregando dataset de dados e ajustando ele
fluxo = pd.read_excel(
    "G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\Fluxo.xlsx")
fluxo.drop('Unnamed: 0', axis=1, inplace=True)

fluxo["Net"] = fluxo["Entrada"] + fluxo["Saida"]
fluxo["Acumulado"] = fluxo["Net"].cumsum()


def menu_admin(nome):
    if nome == "admin":
        c = 0
        while c == 0:
            escolha = int(input("""
            Escolha uma das opções abaixo (digite o numero desejado)
            1 - Total Entrada
            2 - Total Saida
            3 - Entradas por dia (gráfico)
            4 - Saidas por dia (gráfico)
            5 - Entradas/Saidas/Net por dia (gráfico)
            6 - Enviar por email os dados acima
            0 - Sair do app
            """))

            if escolha == 0:
                print("Obrigado por utilizar!!")
                exit()

            if escolha == 1:
                fluxo_entrada = fluxo["Entrada"].sum()
                print(f" Total de Entradas acumulado foi de R${fluxo_entrada:,.2f}")
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 2:
                fluxo_saida = fluxo["Saida"].sum()
                print(f"Total de Saidas acumulado foi de R${fluxo_saida:,.2f}")
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 3:
                x = fluxo["Data"]
                y = fluxo["Entrada"]
                plt.rcParams['xtick.labelsize'] = 8
                plt.rcParams['ytick.labelsize'] = 8
                plt.plot(x, y)
                plt.xticks(rotation=45)
                plt.show()
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 4:
                x = fluxo["Data"]
                y = fluxo["Saida"]
                plt.rcParams['xtick.labelsize'] = 8
                plt.rcParams['ytick.labelsize'] = 8
                plt.plot(x, y, color="Red")
                plt.xticks(rotation=45)
                plt.show()
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 5:
                x = fluxo["Data"]
                y1 = fluxo["Entrada"]
                y2 = fluxo["Saida"]
                y3 = fluxo["Net"]
                plt.plot(x, y1)
                plt.plot(x, y2, color="Red")
                plt.plot(x, y3, color="Black")
                plt.xticks(rotation=45)
                plt.show()
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 6:
                fluxo_entrada = fluxo["Entrada"].sum()

                fluxo_saida = fluxo["Saida"].sum()

                fluxo_net = fluxo["Net"].sum()

                # Entradas por dia (Gráfico)
                x1 = fluxo["Data"]
                y1 = fluxo["Entrada"]
                plt.rcParams['xtick.labelsize'] = 8
                plt.rcParams['ytick.labelsize'] = 8
                plt.plot(x1, y1, color="Blue")
                plt.xticks(rotation=45)
                plt.savefig("entrada.png")

                # Saidas por dia (Gráfico)
                x2 = fluxo["Data"]
                y2 = fluxo["Saida"]
                plt.rcParams['xtick.labelsize'] = 8
                plt.rcParams['ytick.labelsize'] = 8
                plt.plot(x2, y2, color="Red")
                plt.xticks(rotation=45)
                plt.savefig("saida.png")

                # Entrada,Saida,Net por dia (Gráfico)
                x3 = fluxo["Data"]
                y3 = fluxo["Entrada"]
                y3a = fluxo["Saida"]
                y3b = fluxo["Net"]
                plt.plot(x3, y3)
                plt.plot(x3, y3a, color="Red")
                plt.plot(x3, y3b, color="Black")
                plt.xticks(rotation=45)
                plt.savefig("entrada-saida.png")

                # Email
                # criando execução do outlook
                outlook = win32.Dispatch('outlook.application')
                # criar um email "gate" para envio
                email = outlook.CreateItem(0)
                # configurar o email para envio
                email.To = "alex.ssxargemi@gmail.com"
                email.subject = f"Fluxo de Caixa {date.today()}"
                email.HTMLBody = f"Resumo: <br> Entrada : R${fluxo_entrada:,.2f} <br> Saida - R${fluxo_saida:,.2f} <br> Net R${fluxo_net:,.2f} <br> Anexo arquivos complementares compostos por gráfico e tabelas utilizadas de dados <br> <b> Abs </b> "
                anexo1 = "G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\entrada.png"
                anexo2 = "G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\saida.png"
                anexo3 = "G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\entrada-saida.png"
                anexo4 = "G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\Fluxo.xlsx"

                email.Attachments.Add(anexo1)
                email.Attachments.Add(anexo2)
                email.Attachments.Add(anexo3)
                email.Attachments.Add(anexo4)
                email.Send()
                sleep(3)
                print("Enviado com sucesso!!")

                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1


def menu_betas(nome, escolha=0):
    if nome == "betas":
        c = 0
        while c == 0:
            escolha = int(input("""
            Escolha uma das opções abaixo (digite o numero desejado)
            1 - Total Entrada
            2 - Total Saida
            0 - Sair do App
            """))

            if escolha == 0:
                print("Obrigado por utilizar!!")
                exit()

            if escolha == 1:
                fluxo_entrada = fluxo["Entrada"].sum()
                print(f" Total de Entradas acumulado foi de R${fluxo_entrada:,.2f}")
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 2:
                fluxo_saida = fluxo["Saida"].sum()
                print(f"Total de Saidas acumulado foi de R${fluxo_saida:,.2f}")
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1


def menu_manager(nome, escolha=0):
    if nome == "manager":
        c = 0
        while c == 0:
            escolha = int(input("""
            Escolha uma das opções abaixo (digite o numero desejado)
            1 - Total Entrada
            2 - Total Saida
            3 - Entradas por dia (gráfico)
            4 - Saidas por dia (gráfico)
            0 - Sair do app
            """))

            if escolha == 0:
                print("Obrigado por utilizar!!")
                exit()

            if escolha == 1:
                fluxo_entrada = fluxo["Entrada"].sum()
                print(f" Total de Entradas acumulado foi de R${fluxo_entrada:,.2f}")
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 2:
                fluxo_saida = fluxo["Saida"].sum()
                print(f"Total de Saidas acumulado foi de R${fluxo_saida:,.2f}")
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 3:
                x = fluxo["Data"]
                y = fluxo["Entrada"]
                plt.rcParams['xtick.labelsize'] = 8
                plt.rcParams['ytick.labelsize'] = 8
                plt.plot(x, y)
                plt.xticks(rotation=45)
                plt.show()
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
            elif escolha == 4:
                x = fluxo["Data"]
                y = fluxo["Saida"]
                plt.rcParams['xtick.labelsize'] = 8
                plt.rcParams['ytick.labelsize'] = 8
                plt.plot(x, y, color="Red")
                plt.xticks(rotation=45)
                plt.show()
                menu = int(input("Deseja continuar? 1- Sim/ 0 -Não"))
                if menu == 1:
                    continue
                else:
                    c += 1
