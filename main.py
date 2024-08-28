import os
import requests
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv
import json

# Carrega as variáveis de ambiente (apenas para testes locais)
load_dotenv()

# Função para baixar o arquivo XLSX
def baixar_arquivo_xlsx(url, nome_arquivo):
    response = requests.get(url)
    with open(nome_arquivo, 'wb') as file:
        file.write(response.content)
    print(f"Arquivo {nome_arquivo} baixado com sucesso.")

# Função para atualizar Google Sheets
def atualizar_google_sheet(nome_arquivo, sheet_name, sheet_id, credenciais_json):
    # Cria o fluxo OAuth 2.0
    flow = InstalledAppFlow.from_client_config(
        json.loads(credenciais_json),
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )

    # Realiza a autenticação e obtém as credenciais
    creds = flow.run_local_server(port=0)

    # Construa o serviço
    service = build('sheets', 'v4', credentials=creds)

    # Abre a planilha pelo ID
    sheet = service.spreadsheets()

    # Carrega os dados do XLSX usando Pandas
    df = pd.read_excel(nome_arquivo)

    # Converte o DataFrame para lista de listas
    dados = [df.columns.values.tolist()] + df.values.tolist()

    # Atualiza a Google Sheet com os dados
    body = {
        'values': dados
    }
    sheet.values().update(
        spreadsheetId=sheet_id,
        range=sheet_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    print("Google Sheet atualizada com sucesso.")

# Função principal que combina o download e atualização
def executar_tarefa():
    # Carrega as variáveis de ambiente
    url_arquivo = os.getenv("URL_ARQUIVO_XLSX")
    nome_arquivo = "arquivo.xlsx"
    sheet_name = os.getenv("SHEET_NAME")
    sheet_id = os.getenv("SHEET_ID")
    credenciais_json = os.getenv("GOOGLE_CREDENTIALS_JSON")  # Carrega o JSON das credenciais como string

    # Executa o download e atualização da planilha
    baixar_arquivo_xlsx(url_arquivo, nome_arquivo)
    atualizar_google_sheet(nome_arquivo, sheet_name, sheet_id, credenciais_json)

# Executa a tarefa uma vez
executar_tarefa()
