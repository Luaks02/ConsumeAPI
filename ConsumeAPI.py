#from openpyxl import workbook, load_workbook
#import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")

driver = webdriver.Chrome(options=options)
driver.get("https://www.aldo.com.br/categoria/energia-solar?filtro=143")

driver.implicitly_wait(5)

rowshoplist = driver.find_element(By.XPATH, "//div[@class='row shoplist']/ul/*")
rowshoplist.click()
description = driver.find_element(By.XPATH, "//div[@id='description']/strong [contains(text(), 'é composto por')]/following-sibling::br[2]")

print(description.text)



driver.quit()

#Colhendo informações


#wb = load_workbook(filename="Smart_Haus.xlsm", read_only=False,keep_vba=True)
#ws = wb["Dados_Tecnico-Projeto"]

#ws["A1"].value = "Teste"

#wb.save("Smart_Haus.xlsm")


