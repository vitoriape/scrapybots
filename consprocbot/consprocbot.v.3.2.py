import pyautogui
import pyautogui as py_timeout
import pyautogui as py_edge
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from msedge.selenium_tools import EdgeOptions
from msedge.selenium_tools import Edge
from openpyxl import load_workbook
import xlsxwriter as xw
import pandas as pd
import os


cp_workbook_path = 'Z:\\localdoarquivo\\ProcessosConsproc.xlsx'
cp_workbook = load_workbook(cp_workbook_path)

cp_sheet_processos = cp_workbook['Processos']
cp_sheet_andamento = cp_workbook['Andamento']


edge_options = EdgeOptions()
edge_options.use_chromium = True
py_edge = Edge(executable_path='C:\\localdoarquivo\\msedgedriver.exe', options=edge_options)


py_edge.get('https://www.riodasostras.rj.gov.br/consproc/cons_proc1.php')
py_timeout.sleep(2)
    
    
for rProcessos in range(2, len(cp_sheet_processos['A']) + 1):

    novo_nroproc = cp_sheet_processos['A%s' % rProcessos].value
    novo_exercproc = cp_sheet_processos['B%s' % rProcessos].value
    
    py_edge.find_element(By.NAME, "nro_proc").send_keys(novo_nroproc)
    py_timeout.sleep(1)
    py_edge.find_element(By.NAME, "exerc_proc").send_keys(novo_exercproc)
    py_timeout.sleep(1)
    py_edge.find_element(By.NAME, "btnSubmit").send_keys(Keys.RETURN)
    py_timeout.sleep(1)
    
    numberp = py_edge.find_elements(By.XPATH, '/html/body/div[1]/section[4]/div/div/div[2]/div/div[1]/div/table/tbody/tr[1]/td[2]')[0].text
    updatep = py_edge.find_elements(By.XPATH, '/html/body/div[1]/section[4]/div/div/div[2]/div/div[1]/div/table/tbody/tr[2]/td[2]')[0].text
    datep = py_edge.find_elements(By.XPATH, '/html/body/div[1]/section[4]/div/div/div[2]/div/div[1]/div/table/tbody/tr[3]/td[2]')[0].text
    statusp = py_edge.find_elements(By.XPATH, '/html/body/div[1]/section[4]/div/div/div[2]/div/div[1]/div/table/tbody/tr[4]/td/span/strong')[0].text

    
    rAndamento = len(cp_sheet_andamento['A']) + 1
    
    columnA = "A" + str(rAndamento)
    columnB = "B" + str(rAndamento)
    columnC = "C" + str(rAndamento)
    columnD = "D" + str(rAndamento)

    cp_sheet_andamento[columnA] = numberp
    cp_sheet_andamento[columnB] = updatep
    cp_sheet_andamento[columnC] = datep
    cp_sheet_andamento[columnD] = statusp
    
    py_timeout.sleep(2)
    py_edge.find_element(By.XPATH, '/html/body/div[1]/section[4]/div/div/div[2]/div/div[1]/div/div/button').send_keys(Keys.RETURN)
    

py_edge.close()
cp_workbook.save(filename=cp_workbook_path)


cp_user_menu = pyautogui.confirm('Gostaria de abrir o arquivo?', buttons = ['Sim', 'Não'])

if cp_user_menu == 'Sim':
    os.startfile(cp_workbook_path)

elif cp_user_menu == 'Não':
    print("Teste")