import pandas as pd
import datetime
from os.path import getmtime
from datetime import date
from pathlib import Path
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from glob import iglob
from shutil import copyfile
import shutil
import time
import os
import openpyxl
import matplotlib.pyplot as plt
import schedule


def login():

    load_dotenv(r'D:\PRojetos\Codes\GFL\Sinclog(RPA)\Reajuste de Frete\dados_login.env')
    
    user = os.getenv('USER')
    senha= os.getenv('PASSWORD')

    navegador = webdriver.Chrome('C:/Users/PedroFerreiraMouraBa/anaconda3/chromedriver')
    navegador.get("http://gfl.sinclog.com.br/login")
    time.sleep(3)

    username = navegador.find_element('xpath', '//*[@id="login"]')
    password = navegador.find_element('xpath', '//*[@id="senha"]')
    time.sleep(2)

    username.send_keys(user)
    time.sleep(2)

    password.send_keys(senha)
    time.sleep(2)

    navegador.find_element('xpath', '//*[@id="formLogin"]/button').click()
    time.sleep(10)
    return (navegador)


def mover(data):
    files = iglob(r'C:/Users/User\Downloads//entregas*.csv')
    sorted_files = sorted(files, key=getmtime, reverse=True)

    for f in sorted_files:    
        shutil.move(f, 'C:\Base\V2\\Relatorio Geral ' + data + '.csv')
        


def v2(z):
    
    navegador = login()
    navegador.find_element('xpath', '//*[@id="menu_4"]/a/span').click()
    time.sleep(3)
    navegador.find_element('xpath', '//*[@id="menu_4"]/ul/li[7]/a').click()
    time.sleep(3)
    navegador.find_element('xpath', '//*[@id="menu_4"]/ul/li[7]/ul/li/a').click()
    time.sleep(10)
    str_data = (date.today() - datetime.timedelta(days=z)).strftime('%d/%m/%Y')
    DataInicio = navegador.find_element('xpath', '//*[@id="dtIniSolicitacao"]')
    DataInicio.send_keys(str_data)
    time.sleep(2)
    DataFim = navegador.find_element('xpath', '//*[@id="dtFimSolicitacao"]')
    DataFim.send_keys(str_data)
    time.sleep(2)
    navegador.find_element('xpath', '//*[@id="valorBusca"]').click()
    Tipo = navegador.find_element('xpath', '//*[@id="tipoSaida"]')
    Tipo.send_keys("Excel")
    time.sleep(2)
    navegador.find_element('xpath', '/html/body/div[2]/div/section[2]/div[2]/div[2]/form/div[2]/div[2]/div/div[5]/div[1]/button').click()
    navegador.find_element('xpath', '/html/body/div[2]/div/section[2]/div[2]/div[2]/form/div[2]/div[2]/div/div[5]/div[1]/ul/li[2]/a/label/input').click()
    actions = ActionChains(navegador)
    actions.move_to_element(navegador.find_element('xpath', '/html/body/div[2]/div/section[2]/div[2]/div[2]/form/div[2]/div[2]/div/div[6]/div/span')
    ).click().perform()
    navegador.find_element('xpath', '/html/body/div[2]/div/section[2]/div[2]/div[2]/form/div[2]/div[2]/div/div[15]/input[2]').click()
    timeout_relatorio = 5200
    timeout_download = 120
    WebDriverWait(navegador, timeout_relatorio).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btDownload"]'))).click()
    time.sleep(timeout_download)
    navegador.close()

    mover(str_data)


def remu(z):
    

    navegador = login()

    navegador.find_element('xpath', '//*[@id="menu_6"]/a/span').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="menu_6"]/ul/li[5]').click()
    time.sleep(5)   
    navegador.find_element('xpath', '//*[@id="menu_6"]/ul/li[5]/ul/li[2]/a').click()
    time.sleep(5)
    Cadastro_Inicio = navegador.find_element('xpath', '//*[@id="filtros_dtIniCadastro"]')
    Cadastro_Fim = navegador.find_element('xpath', '//*[@id="filtros_dtFimCadastro"]')
    Cadastro_Inicio.clear()
    time.sleep(1)
    Cadastro_Fim.clear()
    time.sleep(1)
    Tipo = navegador.find_element('xpath', '//*[@id="filtros_idPesoConsiderar"]')
    Tipo.send_keys("Peso ou cubagem aferida pelo transportador (o maior)")

    str_data = (date.today() - datetime.timedelta(days=z)).strftime('%d/%m/%Y')

    DataInicio = navegador.find_element('xpath', '//*[@id="filtros_dtIniCadastro"]')
    DataInicio.send_keys(str_data)
    time.sleep(2)
    DataFim = navegador.find_element('xpath', '//*[@id="filtros_dtFimCadastro"]')
    DataFim.send_keys(str_data)
    time.sleep(2)
    Tipo = navegador.find_element('xpath', '//*[@id="form0"]/div/div[1]/div/div[2]').click()

    #selecao do servico 
    navegador.find_element('xpath', '//*[@id="form0"]/div/div[2]/div/div[3]/div/button').click()
    navegador.find_element('xpath', '//*[@id="form0"]/div/div[2]/div/div[3]/div/ul/li[8]/a/label/input').click()
    navegador.find_element('xpath', '//*[@id="form0"]/div/div[2]/div/div[3]/div/ul/li[9]/a/label/input').click()
    navegador.find_element('xpath', '//*[@id="form0"]/div/div[2]/div/div[3]/div/ul/li[2]/a/label/input').click()
    navegador.find_element('xpath', '//*[@id="form0"]/div/div[1]/div/div[2]').click()

    # Selecao do cliente
    navegador.find_element('xpath', '//*[@id="form0"]/div/div[2]/div/div[1]/div[1]/button').click()

    checkbox = navegador.find_element('xpath','//label[contains(text(), "MAGAZINE LUIZA")]/input[@type="checkbox"]')
    if not checkbox.is_selected():
        checkbox.click()

    navegador.find_element('xpath', '//*[@id="form0"]/div/div[2]/div/div[11]/button').click()
    time.sleep(5)     
    timeout_relatorio = 7200
    #aguarda ate 10min para o download do relatorio
    timeout_download = 120  
    WebDriverWait(navegador, timeout_relatorio).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btDownload"]'))).click()
    time.sleep(timeout_download)
    navegador.close()

    mover(str_data)


def main ():


    schedule.every().day.at("09:22").do(remu,7)
    schedule.every().day.at("09:23").do(v2,7)

    while True:
        schedule.run_pending()
        time.sleep(1)




if __name__ == '__main__':
    main()