from unittest import BaseTestSuite
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pandas as pd
import time

nav = webdriver.Chrome()

tabelaProdutos = pd.read_excel("wishlist.xlsx")

nav.get("https://www.google.com/")

produto = 'iphone 12 64 gb'
produto = produto.lower()
termos_banidos = 'mini watch'
termos_banidos = termos_banidos.lower()

lista_termos_banidos = termos_banidos.split(" ")
lsita_termos_produto = produto.split(" ")

preco_minimo = 3000
preco_maximo = 3500

# Pesquisa o nome do produto no Google
nav.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto)
nav.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

# Clica na aba shopping
elementos = nav.find_elements(By.CLASS_NAME, 'hdtb-mitem')
for item in elementos:
  if "Shopping" in item.text:
    item.click()
    break

# Captura dados do produto encontrado
lista_resultados = nav.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')

for resultado in lista_resultados:
  nome = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text
  nome = nome.lower()

  # Verificação do nome do produto
  tem_termos_banidos = False
  for palavra in lista_termos_banidos:
    if palavra in nome:
      tem_termos_banidos = True

  tem_todos_termos_produtos = True
  for palavra in lsita_termos_produto:
    if palavra not in nome:
      tem_todos_termos_produtos = False

  # Verifica se no resultado há termos banidos e se todos os termos do produto
  # estão contidos no nome
  if not tem_termos_banidos and tem_todos_termos_produtos:
    preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
    preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    preco = float(preco)

    # Verifica se preço está dentro da margem aceitável
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    if preco_minimo <= preco <= preco_maximo:
      elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
      elemento_pai = elemento_link.find_element(By.XPATH, '..')
      link = elemento_pai.get_attribute('href')

      print(preco, nome, link)
