import pandas as pd
import numpy as np
from openpyxl import load_workbook

def vertical_merge(datafile_ref, sheetlist):    
    df_list = pd.read_excel(datafile_ref, sheetlist, header=None)

    df_dest = pd.concat(df_list)

    df_dest.to_excel(datafile_ref, sheetlist[0], header=False, index=False)

    del sheetlist[0]
    for sheet in sheetlist:
        del datafile_ref.book[sheet]


def horizontal_merge(datafile_ref, sheetlist):
    df_list = pd.read_excel(datafile_ref, sheetlist, header=None)

    df_dest = pd.concat(df_list, axis=1)

    df_dest.to_excel(datafile_ref, sheetlist[0], header=False, index=False)

    del sheetlist[0]
    for sheet in sheetlist:
        del datafile_ref.book[sheet]


datafile = "Morbidity 1-42.xlsx"
destsheet = "Non-Neonatal Tetanus"
nextsheet = "Table 109"

datafile_ref = pd.ExcelWriter(datafile, mode='a', if_sheet_exists='replace')

vertical_merge(datafile_ref, [destsheet, "Table 107", "Table 108"])
vertical_merge(datafile_ref, [nextsheet, "Table 110", "Table 111"])
horizontal_merge(datafile_ref, [destsheet, nextsheet])

datafile_ref.close()