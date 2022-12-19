#import xlwings as xw

#Abrindo planilha
#wb = xw.Book("smart_haus.xlsm")
#sheet = wb.sheets["Dados_Tecnico-Projeto"]

#sheet.range("A1").value = "Teste"

#wb.save("smart_haus.xlsm")

from openpyxl import workbook, load_workbook

wb = load_workbook(filename="Smart_Haus.xlsm", read_only=False,keep_vba=True)
ws = wb["Dados_Tecnico-Projeto"]

ws["A1"].value = "Teste"

wb.save("Smart_Haus.xlsm")


