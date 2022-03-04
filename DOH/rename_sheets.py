import pandas as pd
import openpyxl

infile = "Morbidity 2018.xlsx"

lookup_df = pd.read_excel("rename_lookuptable.xlsx")
target_wb = openpyxl.load_workbook(infile)

lookup_dict = lookup_df.set_index("OLDNAME").to_dict()["NEWNAME"]

print("Files loaded. Working now...")

for sheet in target_wb:
    if sheet.title in lookup_df["OLDNAME"].values:
        sheet.insert_rows(1)
        sheet["A1"] = lookup_dict[sheet.title]
        sheet.title = lookup_dict[sheet.title][:31]

print("Saving")

target_wb.save(infile)
target_wb.close()