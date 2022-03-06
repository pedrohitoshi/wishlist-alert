# wishlist-alert
Script para consultar os preços da lista de desejos de forma automática

Através da planilha "wishlist.xlsx" o robô pesquisa pelo preço dos produtos listados no Google Shopping e Buscapé.

Caso o preço do produto esteja dentro da margem de preço definida, é enviado um alerta por email para o usuário, contendo todas as informações do produto encontrado e link de compra.

Para que o envio de email funcione corretamente, é importante ter o aplicativo Outlook instalado e configurado na máquina.

A automação foi desenvolvida com Python e Selenium.
