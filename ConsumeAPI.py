#from openpyxl import workbook, load_workbook
#import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")

driver = webdriver.Chrome(options=options)
driver.get("https://www.aldo.com.br/categoria/energia-solar?filtro=143")


rowshoplist = driver.find_element(By.XPATH, "//div[@class='row shoplist']/ul/*")

print(rowshoplist.tag_name)



driver.quit()

#Colhendo informações


#wb = load_workbook(filename="Smart_Haus.xlsm", read_only=False,keep_vba=True)
#ws = wb["Dados_Tecnico-Projeto"]

#ws["A1"].value = "Teste"

#wb.save("Smart_Haus.xlsm")


