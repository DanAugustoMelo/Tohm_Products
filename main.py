import logging
import logging.handlers
import os
import datetime
import pandas as pd
from openpyxl import load_workbook

import requests

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

    # Adicionando a variável de data e hora atual
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    current_time = datetime.datetime.now().strftime("%H:%M:%S")
    
    # Adicionando a variável indicando que o código rodou
    code_status = "O código rodou"

    # Criando um DataFrame pandas com as informações
    df = pd.DataFrame({"Nome": [code_status], "Data": [current_date], "Hora": [current_time]})
    
    # Verificando se o arquivo Excel já existe
    if os.path.isfile("rodou.xlsx"):
        # Carregando o arquivo Excel existente
        wb = load_workbook("rodou.xlsx")
        ws = wb.active
        # Adicionando uma nova linha com as informações atuais
        ws.append(df.iloc[0].tolist())  # Adicionando apenas a primeira linha do DataFrame
        # Salvando as alterações no arquivo Excel
        wb.save("rodou.xlsx")
    else:
        # Caso o arquivo Excel não exista, criar um novo com o DataFrame
        df.to_excel("rodou.xlsx", index=False)