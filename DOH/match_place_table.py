import pandas as pd
import numpy as np

#   This file reads a pdf-to-excel file of the FHSIS and writes the
#   corresponding values according to the place name. Works only for
#   one (1) table.

#   Filename of pdf-to-excel file.
datafilename = "Deworm 1 p337.xlsx"
outfilename = "OUT_" + datafilename
data_sheetname = "Table 1"

#   Some place names from the FHSIS do not match the PCT table.
#   Make dictionary/hash to automate name swap.
rename_dict = {
    'COTABATO': 'COTABATO (NORTH COTABATO)',
    'NORTH COTABATO': 'COTABATO (NORTH COTABATO)',
    'PROVINCE OF DINAGAT': 'DINAGAT ISLANDS',
    'NORTHERN LEYTE': 'LEYTE',
    'MT. PROVINCE': 'MOUNTAIN PROVINCE',
    'MINDORO OCCIDENTAL': 'OCCIDENTAL MINDORO',
    'MINDORO ORIENTAL': 'ORIENTAL MINDORO',
    'WESTERN SAMAR': 'SAMAR (WESTERN SAMAR)',
    'N C R': 'METRO MANILA',
    'NCR': 'METRO MANILA',
    'C A R': 'CAR',
    'REGION 1': 'ILOCOS',
    'REGION 2': 'CAGAYAN VALLEY',
    'REGION 3': 'CENTRAL LUZON',
    'REGION 4A': 'CALABARZON',
    'REGION 4B': 'MIMAROPA',
    'REGION 5': 'BICOL',
    'REGION 6': 'WESTERN VISAYAS',
    'REGION 7': 'CENTRAL VISAYAS',
    'REGION 8': 'EASTERN VISAYAS',
    'REGION 9': 'ZAMBOANGA PENINSULA',
    'REGION 10': 'NORTHERN MINDANAO',
    'REGION 11': 'DAVAO',
    'REGION 12': 'SOCCSARGEN'
}

#   Flag province names that have same name as local name.
#   Not needed for now? Because FHSIS does not report on municipalities
#   with province-similar name. No conflict with HUC/ICC names.
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
#   Removing blank rows. Replacing some names to match place_table names.
data_table = pd.read_excel(datafilename, data_sheetname)
data_table = data_table[data_table['AREA'].str.len() > 0]
data_table['AREA'] = data_table['AREA'].str.upper()
data_table.replace(rename_dict, inplace=True)
data_table.to_excel('OUT_data table.xlsx')

#   Cannot immediately match data_table to place_table because
#   province names repeat. Pick out province rows in place_table
#   then combine separately. Do same for HUC, ICC, and Region for convenience.
prov_df = place_table[place_table['CATEGORY'] == 'Province']
prov_df = prov_df.join(data_table.set_index('AREA'), on='Province')

city_df = place_table[place_table['CATEGORY'].isin(['HUC', 'ICC'])]
city_df = city_df.join(data_table.set_index('AREA'), on='ALT_LOCAL')

region_df = place_table[place_table['CATEGORY'] == 'Region']
region_df = region_df.join(data_table.set_index('AREA'), on='ALT_LOCAL')

place_table[data_table.columns] = np.nan
place_table.update(prov_df)
place_table.update(city_df)
place_table.update(region_df)
place_table.to_excel(outfilename)

#   ----------
#   ----------
#   ----------
