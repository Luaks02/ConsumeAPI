#from openpyxl import workbook, load_workbook
import requests
from bs4 import BeautifulSoup
#import pandas as pd

site = requests.get("https://www.aldo.com.br/busca/energia-solar/aldo-solar-on-grid/deye-microinversor/gerador-de-energia-solar-deye-micro-inversor-sem-estrutura")
soup = BeautifulSoup(site.content,"html.parser")

testao = []

teste = soup.select("img")
for test in teste:
    src = test.get("src")
    alt = test.get("alt")
    testao.append({"src":src,"alt":alt})

print(testao)


#Colhendo informações


#wb = load_workbook(filename="Smart_Haus.xlsm", read_only=False,keep_vba=True)
#ws = wb["Dados_Tecnico-Projeto"]

#ws["A1"].value = "Teste"

#wb.save("Smart_Haus.xlsm")


