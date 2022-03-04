import pandas as pd
import openpyxl

print("Loading workbook")
wb = pd.ExcelFile("Morbidity 2018.xlsx")

df = pd.DataFrame(wb.sheet_names)
df.to_excel("sheetnames.xlsx")

wb.close()