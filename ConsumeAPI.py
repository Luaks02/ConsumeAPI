from openpyxl import workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

#Abrindo a planilha de controle
wb = load_workbook(filename="Smart_Haus.xlsm", read_only=False,keep_vba=True)
ws = wb["Dados_Tecnico-Projeto"]

site = ws["B1"].value

#Colhendo informações com Selenium

options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")

driver = webdriver.Chrome(options=options)
driver.get(site)

try:
    driver.implicitly_wait(10)
    rowshoplist = driver.find_element(By.XPATH, "//div[@class='row shoplist']/ul/*")
    rowshoplist.click()
    driver.implicitly_wait(5)
    description = driver.find_elements(By.XPATH, "//div[@id='description']")[0].text
except NameError:
    print("Não carregou a página, tentar novamente!")

#Filtrando informações colhidas
listed_description = []

listed_description = description.splitlines()

for index,desc in enumerate(listed_description):
    if desc[0:36] == "O gerador de energia fotovoltaico de":
        inicio = index+1
        break

for index,desc in enumerate(listed_description):
    if desc[0:14] == "Regulamentação":
        fim = index-2
        break

#Enviando informações para a planilha de controle

numero = 3

ws["B2"].value = "Equipamentos"
ws["C2"].value = "Qnt"

for linha in range(inicio,fim):
    material = listed_description[linha].split()
    ws["B" + str(numero)].value = " ".join(material[1:])
    ws["C" + str(numero)].value = int(material[0])
    numero += 1

ws["B" + str(numero)].value = "Instalação"
ws["C" + str(numero)].value = 1

#Fechando sistemas

wb.save("Smart_Haus.xlsm")
driver.quit()


