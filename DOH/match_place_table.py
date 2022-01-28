from msilib.schema import ODBCSourceAttribute
import pandas as pd
import numpy as np

#   This file reads a pdf-to-excel file of the FHSIS and writes the
#   corresponding values according to the place name. Works only for
#   one (1) table.

#   Filename of pdf-to-excel file.
datafilename = "datafile_input.xlsx"
outfilename = "output.xlsx"

#   Some place names from the FHSIS do not match the PCT table.
#   Make dictionary/hash to automate name swap.
rename_dict = {
    'COTABATO': 'COTABATO (NORTH COTABATO)',
    'PROVINCE OF DINAGAT': 'DINAGAT ISLANDS',
    'NORTHERN LEYTE': 'LEYTE',
    'MT. PROVINCE': 'MOUNTAIN PROVINCE',
    'MINDORO OCCIDENTAL': 'OCCIDENTAL MINDORO',
    'MINDORO ORIENTAL': 'ORIENTAL MINDORO',
    'WESTERN SAMAR': 'SAMAR (WESTERN SAMAR)'
}

#   Flag province names that have same name as local name.
samename = [
    'ISABELA',
    'QUIRINO',
    'AURORA',
    'BULACAN',
    'QUEZON',
    'RIZAL',
    'BILIRAN',
    'SARANGANI'
]

#   Loading entire place_table into a DataFrame.
#   Add new columns for variables/indicators then write to an out file.
place_table = pd.read_excel("place_table.xlsx", "Sheet1")
place_table['ALT_LOCAL'] = place_table['ALT_LOCAL'].str.upper()

#   Loading entire data file (pdf-to-excel file of FHSIS) into a DataFrame.
#   Match row place name to place_table (place_table) then add data.
data_table = pd.read_excel(datafilename, "Table 1")
data_table = data_table[data_table['AREA'].str.len() > 0]
data_table.replace(rename_dict, inplace=True)

#   Selecting "province" rows, because province name repeats in place_table.
#   Will do the same with HUCs, etc. Set index to province name and
#   use join(). Store copy of original index and restore later to merge
#   with place_table.
prov_df = place_table[place_table['CATEGORY'] == 'Province']
prov_df = prov_df.join(data_table.set_index('AREA'), on='Province')

local_df = place_table[place_table['CATEGORY'].isin(['HUC', 'ICC', 'Region'])]
local_df = local_df.join(data_table.set_index('AREA'), on='ALT_LOCAL')

place_table[data_table.columns] = np.nan
place_table.update(prov_df)
place_table.update(local_df)
place_table.to_excel(outfilename)

#   Get rows in data_table that match province names. Remove empty rows.
#   Find ALT_LOCAL names that have matches with province names.
#   Error checking. tukayo is disposable.
#tukayo = data_table[data_table['ALT_LOCAL'].isin(prov_df['Province'])]
#tukayo = tukayo[tukayo['ALT_LOCAL'].str.len() > 0]
#tukayo = tukayo[tukayo['ALT_LOCAL'].isin(place_table['ALT_LOCAL'])]
#data_table = data_table[~data_table['ALT_LOCAL'].isin(tukayo['ALT_LOCAL'])]

#data_table.to_excel(outfilename)