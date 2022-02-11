from unittest import BaseTestSuite
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pandas as pd
import time

nav = webdriver.Chrome()

tabelaProdutos = pd.read_excel("wishlist.xlsx")

nav.get("https://www.google.com/")

produto = 'iphone 12 64gb'

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
  preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
  nome = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text

  elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
  elemento_pai = elemento_link.find_element(By.XPATH, '..')
  link = elemento_pai.get_attribute('href')

  print(preco, nome, link)
