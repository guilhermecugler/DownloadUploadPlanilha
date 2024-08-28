import requests
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Função para baixar o arquivo XLSX
def baixar_arquivo_xlsx(url, nome_arquivo):
    response = requests.get(url)
    with open(nome_arquivo, 'wb') as file:
        file.write(response.content)
    print(f"Arquivo {nome_arquivo} baixado com sucesso.")

# Função para atualizar Google Sheets
def atualizar_google_sheet(nome_arquivo, sheet_name, sheet_id):
    # Configurações de autenticação com o Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("caminho/para/suas/credenciais.json", scope)
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

# URL do arquivo XLSX para download
url_arquivo = "URL_DO_ARQUIVO_XLSX"
nome_arquivo = "arquivo.xlsx"

# Informações da Google Sheet
sheet_name = "Nome_da_Sheet"
sheet_id = "ID_DA_PLANILHA"

# Executa o download e atualização da planilha
baixar_arquivo_xlsx(url_arquivo, nome_arquivo)
atualizar_google_sheet(nome_arquivo, sheet_name, sheet_id)
