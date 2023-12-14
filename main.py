from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

import pandas as pd
import time
import os
from assets import usuario, senha

#pegando a pasta
diretorio = os.getcwd()

#caminho do chrome drive na minha maquina
chrome_driver = r'C:\Program Files\chromedriver_win32\chromedriver.exe'

#Opções do navegador Chrome
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.default_directory': diretorio,
    'download.prompt_for_downloads':True,
    'download.directory_upgrade':True,
    'safebrowsing.enabled': True
})


#criando o navegador
driver = webdriver.Chrome(executable_path=chrome_driver, options=options)

notas_fiscais = pd.read_excel('NotasEmitir.xlsx')

#Login
def login():
    user = driver.find_element(By.XPATH, '/html/body/div/form/input[1]')
    password = driver.find_element(By.XPATH, '/html/body/div/form/input[2]')

    if user and password:
        time.sleep(1)
        user.send_keys(str(usuario))
        time.sleep(1)
        password.send_keys(str(senha))
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/div/form/button').click()


#Preencher informações do destinatário
def preencher_notas_destinatário(fantasia, endereco, bairro, municipio, cep,uf, cnpj, desc_estadual ):

    try:
        elementos = driver.find_elements(By.XPATH, '//*[@id="nome"]')

        elementos[0].send_keys(fantasia)
        time.sleep(0.5)

        elementos[1].send_keys(endereco)
        time.sleep(0.5)

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[3]').send_keys(bairro)
        time.sleep(0.5)

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[4]').send_keys(municipio)
        time.sleep(0.5)

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[5]').send_keys(str(cep))
        time.sleep(0.5)

        select = driver.find_element(By.XPATH, '//*[@id="formulario"]/select')
        elemento_select = Select(select)
        elemento_select.select_by_visible_text(str(uf))
        

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[6]').send_keys(str(cnpj))
        time.sleep(0.5)

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[7]').send_keys(str(desc_estadual))
        time.sleep(0.5)
    
    except Exception as error:
        print(f'Erro no elemento {error}')


#Preencher informações da mercadoria
def preencher_notas_mercadoria(desc_prod_serv, quantidade, valor_unit, valor_total):
    try:
        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[8]').send_keys(desc_prod_serv)
        time.sleep(0.5)

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[9]').send_keys(str(quantidade))
        time.sleep(0.5)

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[10]').send_keys(str(valor_unit))
        time.sleep(0.5)

        driver.find_element(By.XPATH, '//*[@id="formulario"]/input[11]').send_keys(str(valor_total))
        time.sleep(0.5)

        driver.find_element(By.CLASS_NAME, 'registerbtn').click()

    except Exception as error:
        print(f'Erro no elemento {error}')


#limpar campos
def limpar_campos():
        inputs = driver.find_elements(By.TAG_NAME, 'input')

        for input in inputs:
            input.clear()

        select = driver.find_element(By.XPATH, '//*[@id="formulario"]/select')
        elemento_select = Select(select)
        elemento_select.select_by_index(0)

   
        
        

def main(): 
    
    #direcionando para outra pagina
    driver.get(diretorio+ r'\login.html')
    driver.maximize_window()

    login()

    for i in range(len(notas_fiscais)):

        # DESTINATÁRIO
        nome_fantasia = notas_fiscais.iloc[i]['Cliente']
        endereco = notas_fiscais.iloc[i]['Endereço']
        bairro = notas_fiscais.iloc[i]['Bairro']
        municipio = notas_fiscais.iloc[i]['Municipio']
        cep = notas_fiscais.iloc[i]['CEP']
        UF = notas_fiscais.iloc[i]['UF']
        cnpj_cpf = notas_fiscais.iloc[i]['CPF/CNPJ']
        inscricao_estadual = notas_fiscais.iloc[i]['Inscricao Estadual']

        preencher_notas_destinatário(nome_fantasia,endereco,bairro,municipio,cep,UF,cnpj_cpf,inscricao_estadual)
        
        # MERCADORIA
        desc_prod_serv = notas_fiscais.iloc[i]['Descrição']
        quantidade = notas_fiscais.iloc[i]['Quantidade']
        valor_unitario = notas_fiscais.iloc[i]['Valor Unitario']
        valor_total = notas_fiscais.iloc[i]['Valor Total']

        preencher_notas_mercadoria(desc_prod_serv,quantidade,valor_unitario,valor_total)

        time.sleep(1)
        limpar_campos()
        

    driver.quit()
       





if __name__ == '__main__':
    main()


