from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pandas as pd
import time
import win32com.client as win32

nav = webdriver.Chrome()

tabelaProdutos = pd.read_excel("wishlist.xlsx")


def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):

  nav.get("https://www.google.com/")

  # Trata os valores que vieram da tabela
  produto = produto.lower()
  termos_banidos = termos_banidos.lower()
  lista_termos_banidos = termos_banidos.split(" ")
  lista_termos_produto = produto.split(" ")
  preco_maximo = float(preco_maximo)
  preco_minimo = float(preco_minimo)

  # Pesquisa pelo nome do produto no Google
  time.sleep(10)
  nav.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto, Keys.ENTER)

  # Clica na aba shopping
  elementos = nav.find_elements(By.CLASS_NAME, 'hdtb-mitem')
  for item in elementos:
    if "Shopping" in item.text:
      item.click()
      break

  # Captura dados do produto encontrado
  lista_resultados = nav.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')

  # Verificar as informações de cada resultado correspondem às condições
  lista_ofertas = []
  for resultado in lista_resultados:
    nome = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text
    nome = nome.lower()

    # Verifica se há algum termo banido no nome do produto
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
      if palavra in nome:
        tem_termos_banidos = True

    # Verifica se o nome tem todos os termos do produto
    tem_todos_termos_produtos = True
    for palavra in lista_termos_produto:
      if palavra not in nome:
        tem_todos_termos_produtos = False

    # Verifica se no resultado há termos banidos e se todos os termos do produto
    # estão contidos no nome
    if not tem_termos_banidos and tem_todos_termos_produtos:
      try:
        preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
        preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
        preco = float(preco)

        # Verifica se preço está dentro da margem aceitável
        if preco_minimo <= preco <= preco_maximo:
          elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
          elemento_pai = elemento_link.find_element(By.XPATH, '..')
          link = elemento_pai.get_attribute('href')
          lista_ofertas.append((nome, preco, link))
      except:
        continue


  return lista_ofertas

def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
  
  # Tratar os valores da função
  preco_maximo = float(preco_maximo)
  preco_minimo = float(preco_minimo)
  produto = produto.lower()
  termos_banidos = termos_banidos.lower()
  lista_termos_banidos = termos_banidos.split(" ")
  lista_termos_produto = produto.split(" ")

  # Entrar no buscape
  nav.get('https://www.buscape.com.br/')

  # Pesquisar pelo produto no buscape
  nav.find_element(By.CLASS_NAME, 'search-bar__text-box').send_keys(produto, Keys.ENTER)

  # Pegar a lista de resultados da busca do buscape
  time.sleep(5)
  lista_resultados = nav.find_elements(By.CLASS_NAME, 'Cell_Content__1630r')

  # Para cada resultado
  lista_ofertas = []
  for resultado in lista_resultados:
    try:
      preco = resultado.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
      nome = resultado.get_attribute('title')
      nome = nome.lower()
      link = resultado.get_attribute('href')
      
      # Verifica se tem termos banidos
      tem_termos_banidos = False
      for palavra in lista_termos_banidos:
        if palavra in nome:
          tem_termos_banidos = True

      # Verifica se o nome tem todos os termos do produto
      tem_todos_termos_produtos = True
      for palavra in lista_termos_produto:
        if palavra not in nome:
          tem_todos_termos_produtos = False

      if not tem_termos_banidos and tem_todos_termos_produtos:
        preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
        preco = float(preco)
        if preco_minimo <= preco <= preco_maximo:
          lista_ofertas.append((nome,preco,link))
    except:
      pass
  
  return lista_ofertas

# Construção da tabela de ofertas encontradas
tabela_ofertas = pd.DataFrame()

for linha in tabelaProdutos.index:
  produto = tabelaProdutos.loc[linha, "Nome"]
  termos_banidos = tabelaProdutos.loc[linha, "Termos banidos"]
  preco_minimo = tabelaProdutos.loc[linha, "Preço mínimo"]
  preco_maximo = tabelaProdutos.loc[linha, "Preço máximo"]

  lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
  if lista_ofertas_google_shopping:
    tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preco', 'link'])
    tabela_ofertas = tabela_ofertas.append(tabela_google_shopping)
  else:
    tabela_google_shopping = None

  lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
  if lista_ofertas_buscape:
    tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
    tabela_ofertas = tabela_ofertas.append(tabela_buscape)
  else:
    tabela_buscape = None

print(tabela_ofertas)

# Exportar para o Excel
tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel("ofertas.xlsx", index=False)

# Verifica se existe algum item dentro da tabela de ofertas
if len(tabela_ofertas.index) > 0:
  # Envia email
  outlook = win32.Dispatch('outlook.application')
  mail = outlook.CreateItem(0)
  mail.To = 'pedrohitoshi@gmail.com'
  mail.Subject = 'Encontrei boas ofertas para a sua lista de desejos'
  mail.HTMLBody = f"""
  <p>Hey!</p>
  <p>Analisei os produtos que você pediu e encontrei boas ofertas!</p>
  <p>Dá uma olhada:</p>
  </br>
  {tabela_ofertas.to_html()}
  </br>
  <p>Att,</p>
  <p>Wishlist Alert</p>
  """
  mail.Send()

nav.quit()


# TODO desenvolver uma rotina de execução
# TODO incluir novos sites de busca
# TODO Tentar usar API do Google Shopping / Buscape e outros marketplaces
