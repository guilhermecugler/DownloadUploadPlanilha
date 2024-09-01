from datetime import datetime
import os
import time
import requests
from bs4 import BeautifulSoup
from requests_toolbelt.multipart.encoder import MultipartEncoder
from dotenv import load_dotenv
import json
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from google.cloud import bigquery

# Carrega as variáveis de ambiente
load_dotenv()

# Variáveis de ambiente
nome_aba_planilha_inventario = os.getenv("NOME_ABA_PLANILHA_INVENTARIO")
nome_aba_planilha_relatorio = os.getenv("NOME_ABA_PLANILHA_RELATORIO")
sheet_id = os.getenv("SHEET_ID")
nome_arquivo_inventario = os.getenv("NOME_ARQUIVO_INVENTARIO")
nome_arquivo_relatorio = os.getenv("NOME_ARQUIVO_RELATORIO")
credenciais_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
email = os.getenv("EMAIL")
senha = os.getenv("SENHA")
habilitar_envio_inventario = os.getenv("HABILITAR_ENVIO_INVENTARIO")
habilitar_envio_relatorio = os.getenv("HABILITAR_ENVIO_RELATORIO")
tempo_maximo_espera_relatorio = int(os.getenv("TEMPO_MAXIMO_ESPERA_RELATORIO", 30)) * 60
sheet_id_relatorio = os.getenv("SHEET_ID_RELATORIO")
sheet_id_inventario = os.getenv("SHEET_ID_INVENTARIO")
# Headers para requisições
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
}

def get_csrf_token(session):
    response = session.get('https://casadomedico.stokki.com.br/pt-br/login', headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        csrf_token = soup.find('input', {'name': '_token'})
        return csrf_token.get('value') if csrf_token else ''
    return ''

def perform_login(session, csrf_token):
    login_data = {'_token': csrf_token, 'email': email, 'password': senha}
    response = session.post('https://casadomedico.stokki.com.br/pt-br/login', headers=headers, data=login_data)
    return response.status_code == 200

def download_inventory(session, csrf_token):
    response = session.post(
        'https://casadomedico.stokki.com.br/pt-br/administrator/inventory/excel',
        headers=headers,
        files={'_token': (None, csrf_token)},
    )
    if response.status_code == 200:
        with open(nome_arquivo_inventario, 'wb') as file:
            file.write(response.content)
        print("Download do inventário concluído com sucesso.")
    else:
        print(f"Erro ao baixar o inventário: {response.status_code}")

def generate_report(session, csrf_token):
    url = 'https://casadomedico.stokki.com.br/pt-br/administrator/report/inventory/generate'
    data = {
        '_token': csrf_token,
        'report_name': 'Relatório de Produtos',
        'action': 'product',
        'type': 'today',
        'provider_id': '',
        'client_id': '1'
    }
    encoder = MultipartEncoder(fields=data)
    headers['Content-Type'] = encoder.content_type
    response = session.post(url, headers=headers, data=encoder)
    return response.json() if response.status_code == 200 else None

def check_report_status(session, csrf_token):

    headers.update({
                    'Accept': '*/*',
                    'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7,ja;q=0.6',
                    'Connection': 'keep-alive',
                    'DNT': '1',
                    'Referer': 'https://casadomedico.stokki.com.br/pt-br/administrator/report/export',
                    'Sec-Fetch-Dest': 'empty',
                    'Sec-Fetch-Mode': 'cors',
                    'Sec-Fetch-Site': 'same-origin',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
                    'X-CSRF-TOKEN': csrf_token,
                    'X-Requested-With': 'XMLHttpRequest',
                    'sec-ch-ua': '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
                    'sec-ch-ua-mobile': '?0',
                    'sec-ch-ua-platform': '"Windows"',
                })

    params = {
        '_': '1725071879349',
    }

    atualizar_pedido = session.get('https://casadomedico.stokki.com.br/api/generate/report', params=params, headers=headers, timeout=30)
    print(f"Atualizar pedido {atualizar_pedido}")

    url = 'https://casadomedico.stokki.com.br/pt-br/administrator/report/export/table'
    headers.update({
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'X-CSRF-TOKEN': csrf_token,
        'X-Requested-With': 'XMLHttpRequest',
    })




    params = {
        'draw': '1',
        'columns[0][data]': 'date',
        'columns[0][name]': '',
        'columns[0][searchable]': 'true',
        'columns[0][orderable]': 'false',
        'columns[0][search][value]': '',
        'columns[0][search][regex]': 'false',
        'columns[1][data]': 'name',
        'columns[1][name]': '',
        'columns[1][searchable]': 'true',
        'columns[1][orderable]': 'true',
        'columns[1][search][value]': '',
        'columns[1][search][regex]': 'false',
        'columns[2][data]': 'period',
        'columns[2][name]': '',
        'columns[2][searchable]': 'true',
        'columns[2][orderable]': 'true',
        'columns[2][search][value]': '',
        'columns[2][search][regex]': 'false',
        'columns[3][data]': 'interval',
        'columns[3][name]': '',
        'columns[3][searchable]': 'true',
        'columns[3][orderable]': 'true',
        'columns[3][search][value]': '',
        'columns[3][search][regex]': 'false',
        'columns[4][data]': 'user',
        'columns[4][name]': '',
        'columns[4][searchable]': 'true',
        'columns[4][orderable]': 'true',
        'columns[4][search][value]': '',
        'columns[4][search][regex]': 'false',
        'columns[5][data]': 'client',
        'columns[5][name]': '',
        'columns[5][searchable]': 'true',
        'columns[5][orderable]': 'true',
        'columns[5][search][value]': '',
        'columns[5][search][regex]': 'false',
        'columns[6][data]': 'provider',
        'columns[6][name]': '',
        'columns[6][searchable]': 'true',
        'columns[6][orderable]': 'true',
        'columns[6][search][value]': '',
        'columns[6][search][regex]': 'false',
        'columns[7][data]': 'state',
        'columns[7][name]': '',
        'columns[7][searchable]': 'true',
        'columns[7][orderable]': 'true',
        'columns[7][search][value]': '',
        'columns[7][search][regex]': 'false',
        'columns[8][data]': 'action',
        'columns[8][name]': '',
        'columns[8][searchable]': 'true',
        'columns[8][orderable]': 'true',
        'columns[8][search][value]': '',
        'columns[8][search][regex]': 'false',
        'order[0][column]': '0',
        'order[0][dir]': 'desc',
        'start': '0',
        'length': '10',
        'search[value]': '',
        'search[regex]': 'false',
        'input_search': '',
    }

    start_time = time.time()
    while time.time() - start_time < tempo_maximo_espera_relatorio:
        response = session.get(url, params=params, headers=headers)
        if response.status_code == 200:
            data = response.json()
            state = data.get('aaData', [{}])[0].get('state', '')
            if 'Disponível' in state:
                print("Relatório disponível.")
                action = data.get('aaData', [{}])[0].get('action', '')
                download_link = action.split('href="')[1].split('"')[0]
                download_report(download_link, session)
                break
            elif 'Em processamento' in state:
                print("Relatório ainda em processamento, verificando novamente em 5 minutos.")
                time.sleep(300)
            elif data.get('error') == "Unauthenticated.":
                print("Reautenticando...")
                csrf_token = get_csrf_token(session)
                headers.update({
                            'Accept': 'application/json, text/javascript, */*; q=0.01',
                            'X-CSRF-TOKEN': csrf_token,
                            'X-Requested-With': 'XMLHttpRequest',
                        })
                if perform_login(session, csrf_token):
                    print("Login realizado com sucesso.")
                else:
                    print("Falha ao realizar login.")
                    break
        elif response.status_code == 401:
            print("Reautenticando...")
            csrf_token = get_csrf_token(session)
            headers.update({
                        'Accept': 'application/json, text/javascript, */*; q=0.01',
                        'X-CSRF-TOKEN': csrf_token,
                        'X-Requested-With': 'XMLHttpRequest',
                    })
            if perform_login(session, csrf_token):
                print("Login realizado com sucesso.")
            else:
                print("Falha ao realizar login.")
                break
        else:
            print(f"Erro {response.status_code}: {response.text}")
            # break

def download_report(link, session):
    response = session.get(link)
    with open(nome_arquivo_relatorio, 'wb') as file:
        file.write(response.content)
    print(f'Relatório baixado como {nome_arquivo_relatorio}')

def enviar_para_bigquery(nome_arquivo, dataset_id, table_id, credenciais_json):
    # Carregar as credenciais do JSON
    creds_dict = json.loads(credenciais_json)

    # Configurar o cliente do BigQuery usando as credenciais
    client = bigquery.Client.from_service_account_info(creds_dict)

    # Carregar o arquivo Excel em um DataFrame pandas
    df = pd.read_excel(nome_arquivo, engine='openpyxl')

    # Configurar a referência da tabela no BigQuery
    table_ref = f"{dataset_id}.{table_id}"

    # Configurar o job de carregamento
    job_config = bigquery.LoadJobConfig(
        write_disposition=bigquery.WriteDisposition.WRITE_APPEND,  # Alternativas: WRITE_TRUNCATE, WRITE_EMPTY
        source_format=bigquery.SourceFormat.PARQUET
    )

    # Carregar o DataFrame para o BigQuery
    job = client.load_table_from_dataframe(df, table_ref, job_config=job_config)
    job.result()  # Esperar o job completar

    print(f"Dados enviados para {table_ref} com sucesso.")



def atualizar_google_sheet(nome_arquivo, sheet_name, sheet_id, credenciais_json):
    creds_dict = json.loads(credenciais_json)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(sheet_name)
    df = pd.read_excel(nome_arquivo, engine='openpyxl')
    dados = [df.columns.values.tolist()] + df.values.tolist()

    sanitized_data = [
        [
            "" if isinstance(value, float) and (value != value or value in (float('inf'), -float('inf'))) else value
            for value in row
        ]
        for row in dados
    ]

    worksheet.clear()
    worksheet.update(sanitized_data)
    print("Google Sheet atualizada com sucesso.")

if __name__ == "__main__":
    with requests.Session() as session:
        csrf_token = get_csrf_token(session)
        if perform_login(session, csrf_token):
            print("Login realizado com sucesso.")
            if habilitar_envio_inventario == 'SIM':
                download_inventory(session, csrf_token)
                # atualizar_google_sheet(nome_arquivo_inventario, nome_aba_planilha_inventario, sheet_id_inventario, credenciais_json)
                # Enviar a primeira planilha
                enviar_para_bigquery(nome_arquivo_inventario, 'sheets-370517.RelatorioCasaDoMedico', 'relatorio_inventario', credenciais_json)


            # if habilitar_envio_relatorio == 'SIM':
            #     generate_report(session, csrf_token)
            #     check_report_status(session, csrf_token)
            #     atualizar_google_sheet(nome_arquivo_relatorio, nome_aba_planilha_relatorio, sheet_id_relatorio, credenciais_json)

        else:
            print("Falha ao realizar login.")
