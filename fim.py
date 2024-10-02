import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import numpy as np
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import glob
import chromedriver_autoinstaller
import datetime
from utils import *

while True:
    try:

        scope = ['https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive"]

        credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
        client = gspread.authorize(credentials)
        filename = 'service_account_cemag.json'
        sa = gspread.service_account(filename)

        sheet = 'Planilha de Início/Fim de Validade'
        #worksheet= input('Nome da aba:')
        worksheet = 'Fim de Validade' # Local para fazer alteração do nome da planilha

        sh1 = sa.open(sheet)
        wks1 = sh1.worksheet(worksheet)

        df = wks1.get()

        tabela = pd.DataFrame(df)

        def iframes(nav):

            iframe_list = nav.find_elements(By.CLASS_NAME,'tab-frame')

            for iframe in range(len(iframe_list)):
                time.sleep(1)
                try: 
                    nav.switch_to.default_content()
                    nav.switch_to.frame(iframe_list[iframe])
                    print(iframe)
                except:
                    pass

        def saida_iframe(nav):
            nav.switch_to.default_content()

        def listar(nav, classe):
            
            lista_menu = nav.find_elements(By.CLASS_NAME, classe)
            
            elementos_menu = []

            for x in range (len(lista_menu)):
                a = lista_menu[x].text
                elementos_menu.append(a)

            test_lista = pd.DataFrame(elementos_menu)
            test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

            return(lista_menu, test_lista)


        try:
            nav = webdriver.Chrome(r"C:\Users\pcp2\robo-saldo\chromedriver_extracted\chromedriver-win32\chromedriver.exe")
        except:
            chrome_driver_path = verificar_chrome_driver()
            nav = webdriver.Chrome()
        time.sleep(1)
        nav.maximize_window()
        time.sleep(1)
        nav.get('http://127.0.0.1/sistema')


        nav.find_element(By.ID, 'username').send_keys('joao marcos') #ti.dev
        time.sleep(2)
        nav.find_element(By.ID, 'password').send_keys('280470')
        time.sleep(1)
        nav.find_element(By.ID, 'submit-login').click() 


        WebDriverWait(nav,20).until(EC.presence_of_element_located((By.ID, 'bt_1892603865')))
        time.sleep(1)
        nav.find_element(By.ID, 'bt_1892603865').click()
        time.sleep(3)
        lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
        time.sleep(1)

        click_producao = test_list.loc[test_list[0] == 'Projeto'].reset_index(drop=True)['index'][0]
        lista_menu[click_producao].click()
        time.sleep(2)

        lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
        time.sleep(1)
        click_producao = test_list.loc[test_list[0] == 'Materiais e Produtos'].reset_index(drop=True)['index'][0]
        lista_menu[click_producao].click()
        time.sleep(8)

        iframes(nav)
        WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')))
        time.sleep(1)
        nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div').click()
        time.sleep(3)
        input_localizar = nav.find_element(By.ID, 'grInputSearch_explorer')
        time.sleep(1.5)

        tabela2 = tabela.copy()
        tabela2 = tabela2.rename(columns={0: 'Código', 1:'Status'})
        tabela2 = tabela2.fillna('')
        tabela2 = tabela2[tabela2['Código'].notnull() & (tabela2['Código'] != '') & (tabela2['Status'] == '')]
        tabela2 = tabela2['Código']
        tabela2 = tabela2[2:]
        # Resetar o índice
        tabela2 = tabela2.reset_index()

        input_localizar.send_keys(tabela2['Código'][0])
        time.sleep(1.5)
        input_localizar.send_keys(Keys.ENTER)
        time.sleep(1.5)
        WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[1]/input')))
        time.sleep(1)
        nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[1]/input').click()
        time.sleep(1.5)
        WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[1]/input')))
        time.sleep(1)
        nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div').click()
        time.sleep(1.5)

        for i in range(len(tabela2)):
            
            linha = tabela2['index'][i]
            
            if tabela2['Código'][i] != '' or tabela2['Código'][i] != None:
                time.sleep(1.5)
                input_localizar.send_keys(Keys.CONTROL + 'a')
                time.sleep(2)
                input_localizar.send_keys(Keys.BACKSPACE)
                time.sleep(2)
                input_localizar.send_keys(tabela2['Código'][i])
                time.sleep(2)
                WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[4]/b')))
                time.sleep(1)
                nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[4]/b').click()
                time.sleep(3)
                WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'//*[@id="gridTitle_explorer_RECURSOETAPA"]')))
                time.sleep(1)
                nav.find_element(By.XPATH,'//*[@id="gridTitle_explorer_RECURSOETAPA"]').click()
                time.sleep(2)
                if i == 0:
                    WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div')))
                    time.sleep(1)
                    nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div').click()
                    time.sleep(2)
                time.sleep(1)
                WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]')))
                time.sleep(1)
                nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]').click()
                time.sleep(2)
                if i == 0:
                    WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')))
                    time.sleep(1)
                    nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div').click()
                    time.sleep(1.5)
                time.sleep(1)
                WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[3]/div/div')))
                time.sleep(1)
                nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[3]/div/div').click()
                time.sleep(1.5)
                
                input_localizar_recursos = nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[3]/input')
                time.sleep(1.5)
                input_localizar_recursos.send_keys(Keys.CONTROL + 'a')
                time.sleep(1.5)
                input_localizar_recursos.send_keys(Keys.BACKSPACE)
                time.sleep(1.5)
                input_localizar_recursos.send_keys(Keys.BACKSPACE)
                input_localizar_recursos.send_keys(Keys.BACKSPACE)

                # --------------------------------------------------------- Tratamento tabelas Inicio-------------------------------------------------------------
                tabela3 = tabela.copy()
                tabela3 = tabela3[2]
                tabela_recursos_inseridos = tabela3[2]
                tabela_ordem = tabela3[5]
                tabela_fim = tabela3[8]
                # tabela_obs = tabela3[11]
                # tabela_qtd = tabela3[14]
                # tabela_deposito = tabela3[17]
                # --------------------------------------------------------- Tratamento tabelas Fim -------------------------------------------------------------
                time.sleep(1.5)
                input_localizar_recursos.send_keys(tabela_recursos_inseridos)
                time.sleep(1.5)
                input_localizar_recursos.send_keys(Keys.ENTER)
                time.sleep(1.5)
                # WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[4]/b')))
                # time.sleep(1.5)
                # nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[4]/b')
                # time.sleep(1.5)

                if i == 0:
                    WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[4]/div/div')))
                    time.sleep(1)
                    nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[4]/div/div').click()
                    time.sleep(1.5)
            
                input_localizar_recursos.click()
            
                # Não entendi o que isso faz #

                # WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[13]/td[11]/div/div')))
                # time.sleep(1)
                # nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[13]/td[11]/div/div').click()
                # time.sleep(1.5)
            
            
                # --------------------------------------------------------- Tratamento tabelas HTML Inicio-------------------------------------------------------------
                table = nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table')
                table_html = table.get_attribute('outerHTML')
                df = pd.read_html(str(table_html))
                df1 = df.copy()
                df1 = df1[0]
                        # Lista dos índices das colunas a serem removidas
                indices_colunas = [0, 2, 4, 6, 8, 10, 12, 14, 16, 18]
                # Remover as colunas
                df1 = df1.drop(df1.columns[indices_colunas], axis=1)
                df1 = df1.rename(columns={1:'Ordem *',3:'Recurso',5:'Quantidade *',7:'UM',9:'Depósito Origem',11:'Início Validade Técnica',13:'Fim Validade Técnica',15:'Classe de Opcionais',17:'Observação'})
                df1 = df1[1:]
                df1 = df1.dropna(subset=['Ordem *'])
                df1 = df1.reset_index(drop=True)
                contagem = len(df1['Ordem *'])
                df1['Recurso'] = df1['Recurso'].fillna(tabela_recursos_inseridos)
                df1 = df1[df1['Recurso'].str.contains(tabela_recursos_inseridos)]
                df1 = df1[df1['Fim Validade Técnica'].isnull()]
                df1.reset_index(inplace=True)
                # --------------------------------------------------------- Tratamento tabelas HTML Fim------------------------------------------------------------
                time.sleep(1.5)
                for i2 in range(len(df1)):
                    indice = df1['index'][i2]
                    time.sleep(1.5)
                    WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="'+ str(indice) +'"]/td[10]/div/div')))
                    time.sleep(1)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[2]/div/div').click()
                    time.sleep(1)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[2]/div/input').send_keys(tabela_ordem)
                    time.sleep(1)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[2]/div/input').send_keys(Keys.TAB)
                    time.sleep(1)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[10]/div/div').click()
                    time.sleep(1.5)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[10]/div/input').send_keys(tabela_fim)
                    time.sleep(1.5)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[10]/div/input').send_keys(Keys.TAB)
                    time.sleep(2)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[11]/div/input').send_keys(Keys.TAB)
                    time.sleep(2)
                    nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[13]/div/textarea').send_keys(Keys.TAB)
                    time.sleep(5)
                    try:
                        while len(nav.find_element(By.XPATH, '//*[@id="'+ str(indice) +'"]/td[13]/div/textarea')) > 0:
                            print('Carregando...')
                    except:
                        print('Carregou')

                wks1.update('B' + str(linha+1), 'Ok')

        time.sleep(1)
        WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[1]')))
        time.sleep(1)
        nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[3]/td[1]/table/tbody/tr[1]/td[1]').click()
        time.sleep(2)
        nav.close()
        break
    
    except:
        nav.close()
        