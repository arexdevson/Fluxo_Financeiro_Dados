import pandas as pd
from Funções import Conexao_Sheets
from time import sleep
from Funções import Menu
import os

# habilitando conexão com o sheets
Conexao_Sheets.main()

# time
sleep(2)

# carregando dataset de acesso e ajustando ele
acesso = pd.read_excel("Acesso.xlsx")
acesso.drop('Unnamed: 0', axis=1, inplace=True)

# para fazer um "procv" entre usuario e dar o match com o seu correspondente da senha preciso transformar em listas
lista_usuario = []
lista_senha = []

for elemento in acesso["Usuario"]:
    lista_usuario.append(elemento)

for elemento in acesso["Senha"]:
    lista_senha.append(elemento)

# controle de acesso

c = 3
while c != 0:
    nome = str(input("Digite seu usuario"))
    if nome not in lista_usuario:
        c -= 1
        if c > 0:
            print(f"Usuario não encontrado, tente novamente, você tem mais {c} tentativa(s) de acesso")
            sleep(3)

        if c == 0:
            print("Acesso Negado!! , reinicie o processo e tente novamente")
            exit()
        else:

            continue
    else:
        senha = int(input("Digite sua senha de 3 digitos"))
        posicao = lista_usuario.index(nome)
        if senha == lista_senha[posicao]:
            print("Acesso Liberado !!!")
            c = 0
            sleep(3)

        else:
            print("Senha não identificada, tente novamente!!")
            exit()
    sleep(3)
    print("Carregando menu....")
    sleep(3)

    try:
        if nome == "admin":
            Menu.menu_admin(nome)
        if nome == "betas":
            Menu.menu_betas(nome)
        if nome == "manager":
            Menu.menu_manager(nome)
        else:
            print("")
    except(ValueError,TypeError):
        print("Digite corretamente")

