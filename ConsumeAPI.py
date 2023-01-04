from openpyxl import workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

#Colhendo informações com Selenium

options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")

driver = webdriver.Chrome(options=options)
driver.get("https://www.aldo.com.br/categoria/energia-solar?filtro=131;3:5~30000")

driver.implicitly_wait(5)

rowshoplist = driver.find_element(By.XPATH, "//div[@class='row shoplist']/ul/*")
rowshoplist.click()
driver.implicitly_wait(5)
description = driver.find_elements(By.XPATH, "//div[@id='description']")[0].text

listed_description = []

listed_description = description.splitlines()

#Enviando informações para a planilha principal

wb = load_workbook(filename="Smart_Haus.xlsm", read_only=False,keep_vba=True)
ws = wb["Dados_Tecnico-Projeto"]

numero = 2

for linha in listed_description:
    ws["B" + str(numero)].value = linha
    numero += 1

#Fechando sistemas

wb.save("Smart_Haus.xlsm")
driver.quit()


