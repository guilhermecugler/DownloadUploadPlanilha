import os
import requests
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

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
    # Configurações de autenticação com o Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credenciais_json, scope)
    client = gspread.authorize(creds)

    # Abre a planilha pelo ID
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(sheet_name)

    # Carrega os dados do XLSX usando Pandas
    df = pd.read_excel(nome_arquivo)

    # Converte o DataFrame para lista de listas
    dados = [df.columns.values.tolist()] + df.values.tolist()

    # Atualiza a Google Sheet com os dados
    worksheet.clear()
    worksheet.update(dados)
    print("Google Sheet atualizada com sucesso.")

# Função principal que combina o download e atualização
def executar_tarefa():
    # Carrega as variáveis de ambiente
    url_arquivo = os.getenv("URL_ARQUIVO_XLSX")
    nome_arquivo = "arquivo.xlsx"
    sheet_name = os.getenv("SHEET_NAME")
    sheet_id = os.getenv("SHEET_ID")
    credenciais_json = os.getenv("GOOGLE_CREDENTIALS_JSON_PATH")

    # Executa o download e atualização da planilha
    baixar_arquivo_xlsx(url_arquivo, nome_arquivo)
    atualizar_google_sheet(nome_arquivo, sheet_name, sheet_id, credenciais_json)

# Executa a tarefa uma vez
executar_tarefa()
