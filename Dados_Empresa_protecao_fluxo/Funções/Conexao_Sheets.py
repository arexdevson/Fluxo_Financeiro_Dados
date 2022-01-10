"""
https://www.youtube.com/watch?v=ZU30e4gkV8g&t=2003s&ab_channel=HashtagPrograma%C3%A7%C3%A3o
1° Passo:Abri o google e digitar (google developer console)
2° Passo: criar um projeto
3° Passo: Pesquiso por (Google Drive Api) clico e ativo
4° Passo: abro o menu que está do lado superior esquerdo - apis e serviços - Painel
5° Passo: clica em Tela de permissão Oauth  - usertype (externo) - coloco nome do app, meu email
6 Passo: só dar ok e no final voltar pro painel
7° Passo: no menu lateral a esquerda - credenciais - criar credenciais - id do cliente oauth - tipo de aplicativo (app para computador)
8° Passo: ele vai gerar seu id de cliente e chave secreta do cliente
8.1° Passo: na tela do Oauth do lado esquerdo, nos clicamos e abrimos nossa api, então no status de publicação nos deixamos ele free para produção
9° Passo: na credencial criada, no lado direito da barra de icones que aparece quando passa o mouse nos baixamos o cliente oauth (gera um json)
10° Passo: renomeei o arquivo para client_secret.json
11° Passo: digito no google (google sheets api python) e abro o site e eseguimos os passos
12° Passo: pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
13° Colar código pra integração (segue abaixo)
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'] = só leitura
SCOPES = ['https://www.googleapis.com/auth/spreadsheets'] = tudo liberado

site:'https://docs.google.com/spreadsheets/d/1CUjXX3d-NJxAYeiHBLzdbfek_YOwngXg31l2N_35RBo/edit#gid=0'
id do site :1CUjXX3d-NJxAYeiHBLzdbfek_YOwngXg31l2N_35RBo

na parte do..
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run

        renomeamos o credentials.json para o nome do arquivo que baixamos json
"""

from __future__ import print_function
import pandas as pd
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1CUjXX3d-NJxAYeiHBLzdbfek_YOwngXg31l2N_35RBo'
SAMPLE_RANGE_NAME = 'Acesso!A:B'


# conexão/integração python e sheets

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\Apoio\client_secret.json',
                SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # conectar com planilha do google
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId='1CUjXX3d-NJxAYeiHBLzdbfek_YOwngXg31l2N_35RBo',
                                    range='Acesso!A:B').execute()
        values = result.get('values', [])
        dado = pd.DataFrame(values, columns=["Usuario", "Senha"])
        dado = dado.drop(0, axis=0)
        pd.DataFrame(dado).to_excel("G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\Acesso.xlsx")

    except HttpError as err:
        print(err)

    try:
        service2 = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet2 = service2.spreadsheets()
        result2 = sheet2.values().get(spreadsheetId='11V4phR-wcXuSPE1IW_fMdlRJIZjTH5hbzzjKcKXm0PY',
                                      range='Info!A:C').execute()
        values2 = result2.get('values', [])
        dado2 = pd.DataFrame(values2, columns=["Data", "Entrada", "Saida"])
        dado2 = dado2.drop(0, axis=0)
        pd.DataFrame(dado2).to_excel("G:\Meu Drive\Registros\Geral\PycharmProjects\Hashtag Programação\YOUTUBE\Projetos\Dados_Empresa_protecao_fluxo\Fluxo.xlsx")

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
