from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from IPython.display import display

pd.set_option('display.max_columns', None)

#Abrir o navegador
navegador = webdriver.Chrome()

#Obter cotações
#Dolar
navegador.get("https://www.google.com.br/")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dolar")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacaoDolar = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacaoDolar)
#Euro
navegador.get("https://www.google.com.br/")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacaoEuro = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacaoEuro)
#Ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacaoOuro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute("value")
cotacaoOuro = cotacaoOuro.replace(",", ".")
print(cotacaoOuro)

navegador.quit()

#Atualizar a base de dados
tabela = pd.read_excel("Produtos.xlsx")
display(tabela)
print(tabela.info())
#Cotação
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacaoDolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacaoEuro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacaoOuro)
#Preço de Compra
tabela["Preço Base Reais"] = tabela["Preço Base Original"] * tabela["Cotação"]
#Preço de Venda
tabela["Preço Final"] = tabela["Preço Base Reais"] * tabela["Margem"]

#Exportar
tabela.to_excel("Produtos Novo.xlsx", index=False)