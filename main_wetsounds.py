from bs4 import BeautifulSoup
import requests
import pandas as pd
import openpyxl
from datetime import datetime
import logging
import logging.handlers
import os

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

try:
    SOME_SECRET = os.environ["SOME_SECRET"]
except KeyError:
    SOME_SECRET = "Token not available!"
    #logger.info("Token not available!")
    #raise

if __name__ == "__main__":
    logger.info(f"Token value: {SOME_SECRET}")

    # Função para extrair informações do site e salvar na planilha
    def extrair_informacoes(link, linha):
        requisicao = requests.get(link)
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

        # Encontrando a próxima linha vazia na planilha
        proxima_linha = sheet.max_row + 1
        
        # Salvando as informações na planilha na próxima linha vazia
        sheet[f'A{proxima_linha}'] = nome_produto.text if nome_produto else ''
        sheet[f'B{proxima_linha}'] = preco_produto.text if preco_produto else ''
        sheet[f'C{proxima_linha}'] = sku_produto.text if sku_produto else ''
        sheet[f'D{proxima_linha}'] = image_link
        
        # Adicionando a data na próxima linha vazia
        sheet[f'E{proxima_linha}'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Lendo os links da planilha Excel
    planilha = pd.read_excel('Input_links_wet_sounds.xlsx')

    # Carregando a planilha de produtos existente
    workbook = openpyxl.load_workbook('02.Output_informacoes_produtos_wet_sounds.xlsx')
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
        extrair_informacoes(link, idx + 2)  # Começa da linha 2
        
    # Salvando a planilha
    workbook.save('Output_informacoes_produtos_wet_sounds.xlsx')