from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import datetime
import os
from pathlib import Path
from time import sleep
import pyodbc

class OrdensCompra:
    def __init__(self):
        self.site_link = "https://developer.automationanywhere.com/challenges/automationanywherelabs-supplychainmanagement.html"
        self.site_link_2 = "https://developer.automationanywhere.com/challenges/AutomationAnywhereLabs-POtrackingLookup.html"
        self.dia = datetime.now().day
        self.mes = datetime.now().month
        self.ano = datetime.now().year

    def abrir_link(self):
        self.driver = webdriver.Chrome()
        self.driver.get(self.site_link)
        self.tab = self.driver.window_handles[0]
        self.driver.execute_script("window.open('{}', '_blank');".format(self.site_link_2))

        self.driver.maximize_window() 

    def coletar_po(self):
        self.driver.switch_to.window(self.tab)

        self.lista_po = []

        for i in range(1, 8):
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,  f'//*[@id="PONumber{i}"]')))
            self.po_value = self.driver.find_element(By.XPATH, f'//*[@id="PONumber{i}"]').get_attribute("value")
        
            self.lista_po.append(self.po_value)
                
    def criar_planilha(self, file=None):
        
        if file is None:
            file = f'dados_pedidos_{self.dia}_{self.mes}_{self.ano}.xlsx'

            self.df = pd.DataFrame({'PO': self.lista_po})
            self.df.to_excel(file, index=False, engine='openpyxl')
             
    def inserir_po_extracao(self):
        self.driver.switch_to.window(self.driver.window_handles[1])

        self.states = []
        self.dates = []
        self.orders = []

        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="onetrust-accept-btn-handler"]')))
        self.driver.find_element(By.XPATH, '//*[@id="onetrust-accept-btn-handler"]').click()

        for self.i, self.row in self.df.iterrows():

            self.po_number = self.row['PO']

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dtBasicExample_filter"]/label/input')))
            self.search_po = self.driver.find_element(By.XPATH, '//*[@id="dtBasicExample_filter"]/label/input')
            self.search_po.send_keys(self.po_number)
            self.search_po.clear()

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dtBasicExample"]/tbody/tr/td[5]')))
            self.state = self.driver.find_element(By.XPATH, '//*[@id="dtBasicExample"]/tbody/tr/td[5]')
            self.states.append(self.state.text)
            
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dtBasicExample"]/tbody/tr/td[7]')))
            self.date = self.driver.find_element(By.XPATH, '//*[@id="dtBasicExample"]/tbody/tr/td[7]')
            self.dates.append(self.date.text)

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dtBasicExample"]/tbody/tr/td[8]')))
            self.order = self.driver.find_element(By.XPATH, '//*[@id="dtBasicExample"]/tbody/tr/td[8]')
            self.orders.append(self.order.text)

        self.df['State'] = self.states
        self.df['Ship_Date'] = self.dates
        self.df['Order_Total'] = self.orders
        self.df['FL_800'] = self.df['Order_Total'].replace('[\$,]', '', regex=True).astype(float).apply(lambda x: 1 if x > 800 else 0)

        updated_file = f'dados_pedidos_{self.dia}_{self.mes}_{self.ano}.xlsx'
        self.df.to_excel(updated_file, index=False, engine='openpyxl')
    
    def extrair_agente(self):

        self.driver.switch_to.window(self.tab)
        self.df = pd.read_excel(f'dados_pedidos_{self.dia}_{self.mes}_{self.ano}.xlsx')
        self.file_path = Path.home()/"Downloads/StateAssignments.xlsx"

        self.names = []

        if os.path.exists(self.file_path):
            self.df_dados = pd.read_excel(self.file_path)

        else:
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//html/body/div[1]/div/div[2]/a')))
            self.driver.find_element(By.XPATH, '//html/body/div[1]/div/div[2]/a').click()
            sleep(3)

            self.df_dados = pd.read_excel(self.file_path)

        for state in self.states:
            self.name = self.df_dados.loc[self.df_dados['State'] == state, 'Full Name'] 
            self.names.append(self.name.iloc[0])

        self.df['Agent'] = self.names

        updated_file = f'dados_pedidos_{self.dia}_{self.mes}_{self.ano}.xlsx'
        self.df.to_excel(updated_file, index=False, engine='openpyxl')
    
    def insercao_dados(self):

        self.df = pd.read_excel(f'dados_pedidos_{self.dia}_{self.mes}_{self.ano}.xlsx')

        for i, row in self.df.iterrows():

            self.date = row['Ship_Date']
            self.order = row['Order_Total'].replace('$','')
            self.agent = row['Agent']

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="shipDate{i+1}"]')))
            self.driver.find_element(By.XPATH, f'//*[@id="shipDate{i+1}"]').send_keys(self.date)

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="orderTotal{i+1}"]')))
            self.driver.find_element(By.XPATH, f'//*[@id="orderTotal{i+1}"]').send_keys(self.order)

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="agent{i+1}"]')))
            self.driver.find_element(By.XPATH, f'//*[@id="agent{i+1}"]').send_keys(self.agent)

    def encerrar_print(self):

        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="submit_button"]')))
        self.driver.find_element(By.XPATH, '//*[@id="submit_button"]').click()
            
        screenshot = f'sucesso_{self.dia}_{self.mes}_{self.ano}.png'

        sleep(2)

        self.driver.save_screenshot(screenshot)

    def gerar_banco_dados(self):

        server = 'DESKTOP-N5FGB8S'
        banco = 'BD_POAA'
        tabela = 'TBL_POCHL'

        dados_conexao = (
            f"Driver={{SQL Server}};"
            f"Server={server};"
            f"Database={banco};"
            "Trusted_Connection=yes;"
        )

        conexao = pyodbc.connect(dados_conexao)
        print("Conexão bem sucedida")
        cursor = conexao.cursor()

        for i, linha in self.df.iterrows():
            cursor.execute(f"Insert into {tabela} (PO, STATE, SHIP_DATE, ORDER_TOTAL, AGENT) values (?,?,?,?,?)",
                           linha['PO'], linha['State'], linha['Ship_Date'], linha['Order_Total'], linha['Agent']
        )
            
        cursor.commit()

ordemCompra = OrdensCompra()
ordemCompra.abrir_link()
ordemCompra.coletar_po()
ordemCompra.criar_planilha()
ordemCompra.inserir_po_extracao()
ordemCompra.extrair_agente()
ordemCompra.insercao_dados()
ordemCompra.encerrar_print()
ordemCompra.gerar_banco_dados()



