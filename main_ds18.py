import os
import logging
import logging.handlers
from datetime import datetime

import pyodbc
from bs4 import BeautifulSoup
import requests
import pandas as pd
import openpyxl


# ============================== CONFIGURAÇÃO DO LOGGER ==============================

def configure_logger():
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
    return logger


logger = configure_logger()


# ============================ OBTENÇÃO DO TOKEN SECRETO =============================

def get_github_token():
    try:
        return os.environ["SOME_SECRET"]
    except KeyError:
        logger.info("Token not available!")
        return "Token not available!"


SOME_SECRET = get_github_token()
logger.info(f"Token value: {SOME_SECRET}")


# ======================== FUNÇÃO PARA EXTRAÇÃO DE INFORMAÇÕES ========================

def nova_extracao(link, cursor, sheet):
    try:
        # Web scraping da página do produto
        requisicao = requests.get(link)
        requisicao.raise_for_status()
        site = BeautifulSoup(requisicao.text, "html.parser")

        # Extração das informações do produto
        nome_produto = site.find(class_='product-meta__title heading h3')

        preco_elemento = site.find('span', class_='price--large')
        price_text = preco_elemento.get_text().replace(',', '.').replace('$', '').strip() if preco_elemento else None

        # Verificar se o preço inclui "Sale price" e remover, se presente
        if price_text and price_text.startswith('Sale price'):
            price_text = price_text.replace('Sale price', '').strip()

        # Remover caracteres indesejados e formatar o preço
        if price_text:
            preco_produto = price_text.replace(',', '').replace('.', '').strip()
            # Forçar o tipo de dado para número
            preco_produto = float(preco_produto) if preco_produto else None
        else:
            preco_produto = None

        sku_tag = site.find('span', class_='product-meta__sku-number')
        if sku_tag:
            sku_produto = sku_tag.get_text()
        else:
            sku_produto = None

        div_thumbnail = site.find('div', class_='product__thumbnail')
        if div_thumbnail:
            img_tag = div_thumbnail.find('img')
            if img_tag:
                image_link = img_tag['src']
            else:
                image_link = None
        else:
            image_link = None

        # Inserção dos dados no banco de dados
        sql = """
        INSERT INTO products_price_ds18 (nome_produto, preco_produto, sku_produto, image_link, data_extracao, origem)
        VALUES (?, ?, ?, ?, ?, ?)
        """
        valores = (
            nome_produto.text if nome_produto else '',
            preco_produto,
            sku_produto,
            image_link,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'DS18 USA'
        )
        cursor.execute(sql, valores)

        # Inserção dos dados na planilha Excel
        proxima_linha = sheet.max_row + 1
        sheet[f'A{proxima_linha}'] = nome_produto.text if nome_produto else ''
        sheet[f'B{proxima_linha}'] = preco_produto if preco_produto is not None else ''
        sheet[f'C{proxima_linha}'] = sku_produto if sku_produto else ''
        sheet[f'D{proxima_linha}'] = image_link
        sheet[f'E{proxima_linha}'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        sheet[f'F{proxima_linha}'] = 'DS18 USA'

        # Definindo o formato da célula como número com milhares separadores e duas casas decimais
        if preco_produto:
            sheet[f'B{proxima_linha}'].number_format = '#,##0.00'

    except Exception as e:
        logger.error(f"Erro ao processar o link {link}: {e}")


# ======================= CONFIGURAÇÃO DA CONEXÃO COM O SQL AZURE ======================

def connect_to_azure_sql():
    try:
        server = 'webscrapingtohm.database.windows.net'
        database = 'Daily_Scraping_Brands_Prices'
        username = 'admingeral'
        password = 'Tohm@master'
        driver = '{ODBC Driver 17 for SQL Server}'
        conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};PORT=1433;DATABASE={database};UID={username};PWD={password}')
        cursor = conn.cursor()
        return conn, cursor
    except Exception as e:
        logger.error(f"Erro ao conectar ao SQL Azure: {e}")
        return None, None


# ======================== LEITURA E CONFIGURAÇÃO DAS PLANILHAS ========================

def configure_spreadsheets():
    try:
        # Lendo os links da planilha Excel
        planilha = pd.read_excel('03.Input_links_ds18.xlsx')

        # Carregando a planilha de produtos existente
        workbook = openpyxl.load_workbook('02.Output_informacoes_produtos_ds18.xlsx')
        sheet = workbook.active

        # Verificando se as colunas já existem na planilha
        if sheet['A1'].value is None or 'Nome do Produto' not in sheet['A1'].value:
            sheet['A1'] = 'Nome do Produto'
        if sheet['B1'].value is None or 'Preço do Produto' not in sheet['B1'].value:
            sheet['B1'] = 'Preço do Produto'
        if sheet['C1'].value is None or 'SKU do Produto' not in sheet['C1'].value:
            sheet['C1'] = 'SKU do Produto'
        if sheet['D1'].value is None or 'Imagem do Produto' não está em sheet['D1'].value:
            sheet['D1'] = 'Imagem do Produto'
        if sheet['E1'].value is None or 'Data da Extração' não está em sheet['E1'].value:
            sheet['E1'] = 'Data da Extração'
        if sheet['F1'].value is None or 'Origem' não está em sheet['F1'].value:
            sheet['F1'] = 'Origem'

        return planilha, workbook, sheet
    except Exception as e:
        logger.error(f"Erro ao configurar planilhas: {e}")
        return None, None, None


# ================================ EXECUÇÃO DO PROCESSO ================================

if __name__ == "__main__":
    try:
        # Conectando ao SQL Azure
        conn, cursor = connect_to_azure_sql()
        if conn is None or cursor is None:
            raise Exception("Falha ao conectar ao SQL Azure.")

        logger.info("Conexão com SQL Azure estabelecida com sucesso.")

        # Configurando as planilhas
        planilha, workbook, sheet = configure_spreadsheets()
        if planilha is None or workbook is None or sheet is None:
            raise Exception("Falha ao configurar as planilhas.")

        logger.info("Planilhas configuradas com sucesso.")

        # Iterando sobre cada linha da planilha de links e extraindo informações
        for idx, row in planilha.iterrows():
            link = row['Link']
            logger.info(f"Processando link {link}")
            nova_extracao(link, cursor, sheet)

        # Commit para salvar todas as alterações no banco de dados
        conn.commit()
        logger.info("Alterações salvas no banco de dados.")

        # Fechando a conexão
        cursor.close()
        conn.close()
        logger.info("Conexão com SQL Azure fechada.")

        # Salvando a planilha
        workbook.save('02.Output_informacoes_produtos_ds18.xlsx')
        logger.info("Planilha salva com sucesso.")

    except Exception as e:
        logger.error(f"Erro durante a execução do processo: {e}")

    logger.info("Processo concluído com sucesso.")