import pyodbc
from bs4 import BeautifulSoup
import requests
import pandas as pd
import openpyxl
from datetime import datetime
import logging
import logging.handlers
import os

# Configuração do logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
logger_file_handler = logging.handlers.RotatingFileHandler(
    "status.log",
    maxBytes=1024 * 1024,
    backupCount=1,
    encoding="utf8",
)
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger_file_handler.setFormatter(formatter)
logger.addHandler(logger_file_handler)

# Obtenção do token secreto
try:
    SOME_SECRET = os.environ["SOME_SECRET"]
except KeyError:
    SOME_SECRET = "Token not available!"
    #logger.info("Token not available!")

# Função para extrair informações do site e salvar na planilha e no banco de dados
def extrair_informacoes(link, cursor, sheet):
    try:
        requisicao = requests.get(link)
        requisicao.raise_for_status()
        site = BeautifulSoup(requisicao.text, "html.parser")
        
        nome_produto = site.find(class_='productView-title')
        preco_produto = site.find(class_='price price--withoutTax')
        sku_produto = site.find(class_='productView-info')
        image_tag = site.find('img', class_='productView-image--default-custom')

        # Verificar se a tag de imagem foi encontrada
        if image_tag:
            image_link = image_tag.get('data-src', '')  # Usando .get() para evitar erros caso 'data-src' não esteja presente
        else:
            image_link = ''

        # Tratamento do preço
        if preco_produto:
            preco_texto = preco_produto.text.strip()
            if '-' in preco_texto:  # Se houver um intervalo de preços
                preco_texto = preco_texto.split('-')[1].strip()  # Pegar o segundo valor após o "-"
            preco_produto = preco_texto
        else:
            preco_produto = ''

        if preco_produto:
            preco_produto = preco_produto.replace('$', '').replace(',', '').replace('.', ',')
            if ',' in preco_produto and '.' in preco_produto:
                preco_produto = preco_produto.replace(',', '.').replace('.', ',', 1)
            if preco_produto.strip():
                preco_produto = float(preco_produto.replace(',', '.'))

        # Preparar a instrução SQL para inserir os dados
        sql = """
        INSERT INTO produtos_wet_sounds (nome_produto, preco_produto, sku_produto, image_link, data_extracao)
        VALUES (?, ?, ?, ?, ?)
        """
        valores = (
            nome_produto.text if nome_produto else '',
            preco_produto,
            sku_produto.text if sku_produto else '',
            image_link,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
        cursor.execute(sql, valores)

        # Encontrando a próxima linha vazia na planilha
        proxima_linha = sheet.max_row + 1

        # Salvando as informações na planilha na próxima linha vazia
        sheet[f'A{proxima_linha}'] = nome_produto.text if nome_produto else ''
        sheet[f'B{proxima_linha}'] = preco_produto if preco_produto else ''
        sheet[f'C{proxima_linha}'] = sku_produto.text if sku_produto else ''
        sheet[f'D{proxima_linha}'] = image_link
        sheet[f'E{proxima_linha}'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    except Exception as e:
        logger.error(f"Erro ao processar o link {link}: {e}")

if __name__ == "__main__":
    logger.info(f"Token value: {SOME_SECRET}")

    # Configuração da conexão com o SQL Azure
    server = 'webscrapingtohm.database.windows.net'
    database = 'Daily_Scraping_Brands_Prices'
    username = 'admingeral'
    password = 'Tohm@master'
    driver = '{ODBC Driver 17 for SQL Server}'

    # Estabelecendo a conexão
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};PORT=1433;DATABASE={database};UID={username};PWD={password}')
    cursor = conn.cursor()

    # Lendo os links da planilha Excel
    planilha = pd.read_excel('Input_links_wet_sounds2.xlsx')

    # Carregando a planilha de produtos existente
    workbook = openpyxl.load_workbook('Output_informacoes_produtos_wet_sounds2.xlsx')
    sheet = workbook.active

    # Verificando se as colunas já existem na planilha
    if 'Nome do Produto' not in sheet['A1'].value:
        sheet['A1'] = 'Nome do Produto'
    if 'Preço do Produto' not in sheet['B1'].value:
        sheet['B1'] = 'Preço do Produto'
    if 'SKU do Produto' not in sheet['C1'].value:
        sheet['C1'] = 'SKU do Produto'
    if 'Imagem do Produto' not in sheet['D1'].value:
        sheet['D1'] = 'Imagem do Produto'
    if 'Data da Extração' not in sheet['E1'].value:
        sheet['E1'] = 'Data da Extração'

    # Iterando sobre cada linha da planilha de links
    for idx, row in planilha.iterrows():
        link = row['Link']
        extrair_informacoes(link, cursor, sheet)

    # Commit para salvar todas as alterações no banco de dados
    conn.commit()

    # Fechando a conexão
    cursor.close()
    conn.close()

    # Salvando a planilha
    workbook.save('Output_informacoes_produtos_wet_sounds2.xlsx')

    logger.info("Processo concluído com sucesso.")

