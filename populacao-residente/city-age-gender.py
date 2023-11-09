from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
import time
from selenium.common.exceptions import (
    NoSuchElementException,
    ElementNotInteractableException,
)
import os
import shutil
import re




def access_site():
    driver.get("http://tabnet.datasus.gov.br/cgi/deftohtm.exe?ibge/cnv/popsvsbr.def")
    time.sleep(3)

def select_options():
    try:
        driver.find_element(By.XPATH, "//*[@id='L']/option[16]").click() # seleciona Idade Simples na LINHA
        driver.find_element(By.XPATH, "//*[@id='C']/option[6]").click() # seleciona Sexo na COLUNA
        driver.find_element(By.XPATH, "//*[@id='fig4']").click() # seleciona Capitais
    except (NoSuchElementException, ElementNotInteractableException):
        pass

def get_years_options():
    try:
        return driver.find_element(By.XPATH, "//*[@id='A']").find_elements(By.TAG_NAME, "option")
    except (NoSuchElementException, ElementNotInteractableException):
        pass

def get_capital_options():
    try:
        return driver.find_element(By.XPATH, "//*[@id='S4']").find_elements(By.TAG_NAME, "option")
    except (NoSuchElementException, ElementNotInteractableException):
        pass

def create_capital_dir(capital):
    # CRIA DIRETORIO COM NOME DA CAPITAL
    diretorio = os.path.join(os.path.expanduser("~")+"\\Documents", capital.text)
    if not os.path.exists(diretorio):
        os.makedirs(diretorio)

def create_year_dir(capital, ano):
    # CRIA DIRETORIO DE CADA ANO
    diretorio = os.path.join(os.path.expanduser("~")+f"\\Documents\\{capital.text}", ano.text)
    if not os.path.exists(diretorio):
        os.makedirs(diretorio)

def download_csv():
    try:
        Keys.END
        driver.find_element(By.XPATH, "/html/body/div/div/center/div/form/div[4]/div[2]/div[2]/input[1]").click()
        time.sleep(5)
        driver.switch_to.window(driver.window_handles[1])
        Keys.END
        driver.find_element(By.XPATH, "/html/body/div/div/div[3]/table/tbody/tr/td[1]/a").click()
        time.sleep(5)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    except (NoSuchElementException, ElementNotInteractableException):
        pass

def move_csv_file(capital, ano):
    pattern = r'ibge_cnv_popsvsbr(\d+)_\d+_\d+_\d+'
    # Varre os arquivos na pasta de downloads
    for filename in os.listdir(os.path.expanduser("~")+"\\Downloads"):
        match = re.match(pattern, filename)
        if match:
            source_file = os.path.join(os.path.expanduser("~")+"\\Downloads", filename)
            destination_file = os.path.join(os.path.expanduser("~")+f"\\Documents\\{capital.text}\\{ano.text}", filename)
            # Move o arquivo da pasta de downloads para a pasta de destino
            shutil.move(source_file, destination_file)

def next_year():
    driver.find_element(By.XPATH, "//*[@id='A']").send_keys(Keys.ARROW_DOWN)
    time.sleep(2)

def next_capital():
    driver.find_element(By.XPATH, "//*[@id='S4']").send_keys(Keys.ARROW_DOWN)
    time.sleep(3)

if __name__ == "__main__":
    ano = 1
    options = webdriver.EdgeOptions()
    options.add_argument("log-level=3")
    driver = webdriver.Edge(
        service = Service("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedgedriver.exe"),
        options=options)
    opcoes_ano = list()
    opcoes_capital = list()
    try:
        access_site()
        select_options()
        time.sleep(3)
        opcoes_ano = get_years_options()
        opcoes_capital = get_capital_options()
        
        for capital in opcoes_capital:
            if capital.get_attribute("value") == "TODAS_AS_CATEGORIAS__":
                driver.find_element(By.XPATH, "//*[@id='S4']").send_keys(Keys.ARROW_DOWN)
                continue
            
            create_capital_dir(capital)

            # PERCORRE CADA ANO DISPONIVEL POR CAPITAL
            for ano in opcoes_ano:
                create_year_dir(capital, ano)
                download_csv()
                move_csv_file(capital, ano)
                if ano == opcoes_ano[-1]:
                    driver.find_element(By.XPATH, "//*[@id='A']").send_keys(Keys.HOME)
                    break
                next_year()
            next_capital()        

    except (NoSuchElementException, ElementNotInteractableException):
        pass
