from selenium import webdriver
from msedge.selenium_tools import EdgeOptions
from msedge.selenium_tools import Edge
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyautogui
import pyautogui as pt
import pyautogui as nav
from openpyxl import load_workbook
import os


edge_options = EdgeOptions()
edge_options.use_chromium = True
nav = Edge(executable_path=r'C:\Users\PC\Documents\aplicações\msedgedriver.exe', 
            options=edge_options)#Informa o path com o webdriver a ser executado pelo selenium

nav.maximize_window()
nav.get("https://www.mercadolivre.com.br/")#Acessa o site do Mercado Livre

nav.find_element(By.CLASS_NAME, "nav-search-input").send_keys("fogão portátil")#Pesquisa por "fogão portátil" na barra de pesquisa
pt.sleep(1)
nav.find_element(By.CLASS_NAME, "nav-search-btn").click()#Pressiona o botão de pesquisar
pt.sleep(6) 

ads_col = nav.find_elements(By.CLASS_NAME, "ui-search-layout__item")#Identifica todos os anúncios da página pela <li class="">


#Colunas da Planilha: Nome | Valor | URL
ws_workbook_path = r'C:\Users\PC\Downloads\ws_sheet.xlsx'#Caminho da planilha
ws_open_workbook = load_workbook(ws_workbook_path)#Abertura do arquivo

Line = 2

for product in ads_col:

    ws_fogao_sheet = ws_open_workbook['Fogao']#Identificação da guia "Fogao" da planilha ws_sheet.xlsx

    product_image = product.find_element(By.TAG_NAME, "a").get_attribute("href")#Pega a URL do anúncio
    product_name = product.find_element(By.CLASS_NAME, "ui-search-item__title").text#Obtém o nome do produto
    product_value = product.find_element(By.CLASS_NAME, "price-tag-amount").text#Pega o valor do produto sem os centavos


    Line = len(ws_fogao_sheet['A']) + 1#Loop de linhas

    ColumnA = "A" + str(Line)#Coluna A + número da linha
    ColumnB = "B" + str(Line)#Coluna B + número da linha
    ColumnC = "C" + str(Line)#Coluna C + número da linha

    ws_fogao_sheet[ColumnA] = product_name#Indexa o nome do produto na coluna A
    ws_fogao_sheet[ColumnB] = product_value#Indexa o valor concatenado do produto na coluna B
    ws_fogao_sheet[ColumnC] = product_image#Indexa a URL do produto na coluna C


nav.close()#Fecha o navegador
ws_open_workbook.save(filename=ws_workbook_path)#Salva o arquivo

open_menu = pyautogui.confirm("Do you like to open the file?", buttons= ['Sim', 'Não'])
if open_menu == 'Sim':
    os.startfile(ws_workbook_path)

elif open_menu == 'Não':
    print("The info is now saved but you don't wanna see it.")

