#from openpyxl import workbook, load_workbook
import requests
from bs4 import BeautifulSoup
#import pandas as pd

site = requests.get("https://www.aldo.com.br/categoria/energia-solar?filtro=143").text
soup = BeautifulSoup(site,"html.parser")

teste = soup.body.div.div.div.main.div.div


print(teste)





#Colhendo informações


#wb = load_workbook(filename="Smart_Haus.xlsm", read_only=False,keep_vba=True)
#ws = wb["Dados_Tecnico-Projeto"]

#ws["A1"].value = "Teste"

#wb.save("Smart_Haus.xlsm")


