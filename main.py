from datetime import datetime
import glob
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
# Carrega as variáveis de ambiente
load_dotenv()

sheet_name = os.getenv("SHEET_NAME")
sheet_id = os.getenv("SHEET_ID")
nome_arquivo_inventario = os.getenv("NOME_ARQUIVO_INVENTARIO")
nome_arquivo_relatorio = os.getenv("NOME_ARQUIVO_RELATORIO")
sheet_id_relatorio = os.getenv("SHEET_ID_RELATORIO")
sheet_id_inventario = os.getenv("SHEET_ID_INVENTARIO")
credenciais_json = os.getenv("GOOGLE_CREDENTIALS_JSON")  # Carrega o JSON das credenciais como string
email = os.getenv("EMAIL")
senha = os.getenv("SENHA")
# Definindo os headers para as requisições
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7,ja;q=0.6',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'DNT': '1',
    'Referer': 'https://casadomedico.stokki.com.br/pt-br/administrator/inventory',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
}

def get_csrf_token(session):
    response = session.get('https://casadomedico.stokki.com.br/pt-br/login', headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        csrf_token = soup.find('input', {'name': '_token'})
        if csrf_token:
            return csrf_token.get('value')
    return ''

def perform_login(session, csrf_token):
    login_data = {
        '_token': csrf_token,
        'email': email,
        'password': senha,
    }
    login_response = session.post('https://casadomedico.stokki.com.br/pt-br/login', headers=headers, data=login_data)
    return login_response.status_code == 200

def download_inventory(session, csrf_token):
    files = {'_token': (None, csrf_token)}
    response = session.post(
        'https://casadomedico.stokki.com.br/pt-br/administrator/inventory/excel',
        headers=headers,
        files=files,
    )
    if response.status_code == 200:
        with open(nome_arquivo_inventario, 'wb') as file:
            file.write(response.content)
        print("Download do inventário concluído com sucesso.")
    else:
        print(f"Erro ao baixar o inventário: {response.status_code}")

def generate_report(session, csrf_token):

    url = 'https://casadomedico.stokki.com.br/pt-br/administrator/report/inventory/generate'

    # Dados do formulário
    data = {
        '_token': csrf_token,  # Substitua pelo token real
        'report_name': 'Relatório de Produtos',  # Substitua pelos valores reais
        'action': 'product',  # Substitua pela ação real
        'type': 'today',  # Substitua pelo tipo real
        'provider_id': '',
        'client_id': '1'
    }

    # Crie um MultipartEncoder para enviar os dados como multipart/form-data
    encoder = MultipartEncoder(fields=data)

    # Headers da requisição
    headers = {
        'Accept': '*/*',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7,ja;q=0.6',
        'Connection': 'keep-alive',
        'Content-Type': encoder.content_type,
        'DNT': '1',
        'Origin': 'https://casadomedico.stokki.com.br',
        'Referer': 'https://casadomedico.stokki.com.br/pt-br/administrator/report/inventory',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
        'X-CSRF-TOKEN': csrf_token,  # Substitua pelo token real
        'X-Requested-With': 'XMLHttpRequest'
    }


    # Enviar a requisição POST usando a sessão
    response = session.post(url, headers=headers, data=encoder)

    return response.json() if response.status_code == 200 else None

def check_report_status(session, csrf_token):
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
    timeout = 30 * 60  # 30 minutos em segundos

    while True:
        current_time = time.time()
        elapsed_time = current_time - start_time
        formatted_time = datetime.fromtimestamp(current_time).strftime('%H:%M:%S')

        if elapsed_time > timeout:
            print("Tempo limite de 30 minutos alcançado.")
            break

        try:
            response = session.get(url, params=params, headers=headers)
            data = response.json()
            state = data.get('aaData', [{}])[0].get('state', '')
            print(f"Estado atual: {state}")

            if state == '<div class="text-center"><span class="badge badge-success badge-status">Disponível</span></div>':
                print("Relatório disponível.")
                action = data.get('aaData', [{}])[0].get('action', '')
                download_link = action.split('href="')[1].split('"')[0]
                download_report(download_link, session)
                break
            elif state == '<div class="text-center"><span class="badge badge-secondary badge-status">Em processamento</span></div>':
                print(f"Relatório ainda em processamento, verificando novamente em 5 minutos. Hora atual: {formatted_time}")
                time.sleep(300)  # Aguardar 5 minutos
            elif data['error'] == "Unauthenticated.":
                print("Tentando fazer login novamente...")
                csrf_token = get_csrf_token(session)
                if perform_login(session, csrf_token):
                    print("Login realizado com sucesso.")
                else:
                    print("Falha ao realizar login.")
                    break
            else:
                print(response.status_code)
                print(response.text)
                print("Estado desconhecido ou erro na requisição.")
                break
        except requests.RequestException as e:
            print(f"Erro na requisição: {e}")
            print("Tentando fazer login novamente...")
            csrf_token = get_csrf_token(session)
            if perform_login(session, csrf_token):
                print("Login realizado com sucesso.")
            else:
                print("Falha ao realizar login.")
                break


def download_report(link, session):
    response = session.get(link)
    filename = nome_arquivo_relatorio
    
    with open(filename, 'wb') as file:
        file.write(response.content)
    
    print(f'Relatório baixado como {filename}')


def atualizar_google_sheet(nome_arquivo, sheet_name, sheet_id, credenciais_json):
    # Carrega as credenciais do JSON diretamente da string da variável de ambiente
    print("antes")
    print(creds_dict)
    creds_dict = json.loads(credenciais_json)
    print("depois")
    print(creds_dict)
   
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)

    # Abre a planilha pelo ID
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(sheet_name)

    # Carrega os dados do XLSX usando Pandas
    df = pd.read_excel(nome_arquivo, engine='openpyxl')

    # Converte o DataFrame para lista de listas
    dados = [df.columns.values.tolist()] + df.values.tolist()

    print(dados)

    # Atualiza a Google Sheet com os dados
    worksheet.clear()
    worksheet.update(dados)
    print("Google Sheet atualizada com sucesso.")

def delete_all_xlsx_files():
    for file in glob.glob('*.xlsx'):
        os.remove(file)
    print("Arquivos .xlsx excluídos.")

def main():




    session = requests.Session()
    csrf_token = get_csrf_token(session)
    if not csrf_token:
        print("Não foi possível obter o token CSRF.")
        return

    if perform_login(session, csrf_token):
        download_inventory(session, csrf_token)
        # report_response = generate_report(session, csrf_token)

        report_response = True
        if report_response:
            check_report_status(session, csrf_token)
        else:
            print("Falha ao gerar o relatório.")
    else:
        print("Falha no login.")
    atualizar_google_sheet(nome_arquivo_inventario, sheet_name, sheet_id_inventario, credenciais_json)
    atualizar_google_sheet(nome_arquivo_relatorio, sheet_name, sheet_id_relatorio, credenciais_json)
    delete_all_xlsx_files()

if __name__ == '__main__':
    main()
