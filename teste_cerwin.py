{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": [
    "from bs4 import BeautifulSoup\n",
    "import requests\n",
    "import pandas as pd\n",
    "import openpyxl\n",
    "from datetime import datetime\n",
    "\n",
    "link = 'https://cerwinvega.com/spro2100-1d-stroker-pro-2100w-class-d-mono-amplifier.html'\n",
    "# Fazendo a requisição com o proxy\n",
    "requisicao = requests.get(link)\n",
    "\n",
    "# Verificar se a requisição foi bem sucedida\n",
    "if requisicao.status_code == 200:\n",
    "    site = BeautifulSoup(requisicao.text, \"html.parser\")\n",
    "    \n",
    "    nome_produto = site.find(class_='item product')\n",
    "    preco_produto = site.find('span', class_='price')\n",
    "    sku_produto = site.find(class_='value')\n",
    "    image_tag = site.find('img', class_='fotorama__caption')\n",
    "    # Verificar se a tag de imagem foi encontrada\n",
    "    if image_tag:\n",
    "        image_link = image_tag.get('data-src', '')  # Usando .get() para evitar erros caso 'data-src' não esteja presente\n",
    "    else:\n",
    "        image_link = ''\n",
    "\n",
    "    print('Nome do Produto:', nome_produto)\n",
    "    print('Preço:', preco_produto)\n",
    "    print('SKU do Produto:', sku_produto)\n",
    "    print('Link da Imagem:', image_link)\n",
    "else:\n",
    "    print('Erro ao fazer a requisição:', requisicao.status_code)"
   ]
  }
 ],
 "metadata": {
  "language_info": {
   "name": "python"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 2
}
