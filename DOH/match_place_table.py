import pandas as pd
import numpy as np

#   This file reads a pdf-to-excel file of the FHSIS and writes the
#   corresponding values according to the place name. Works only for
#   one (1) table.

#   Filename of pdf-to-excel file.
datafilename = "Morbidity 2019.xlsx"
outfilename = "MATCHED_" + datafilename

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
    'REGION 12': 'SOCCSARGEN',
    'A.R.M.M.': 'BARMM',
    'SOCCSKSARGEN': 'SOCCSARGEN',
    'OLONGAPO CITY': 'OLONGAPO',
    'MEYCAUAYAN CITY': 'MEYCAUAYAN',
    'ILOCOS NOTE': 'ILOCOS NORTE',
    'TAGUIG CITY': 'TAGUIG',
    'NAVOTAS CITY': 'NAVOTAS',
    'MALABON CITY': 'MALABON',
    'SAN JUAN CITY': 'SAN JUAN'
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

datafile_ref = pd.ExcelFile(datafilename)
outfile_ref = pd.ExcelWriter(outfilename)
tempfile_ref = pd.ExcelWriter("OUT_UNMATCHED_DATA.xlsx")

#   Loading entire place_table into a DataFrame.
#   Add new columns for variables/indicators then write to an out file.
place_table = pd.read_excel("place_table.xlsx", "Sheet1")
place_table['ALT_LOCAL'] = place_table['ALT_LOCAL'].str.upper()

#   load entire data workbook into a dataframe list
databook = pd.read_excel(datafile_ref, datafile_ref.sheet_names)

for sheet_name in datafile_ref.sheet_names:
    match_table = place_table.copy()
    #   ALL CAPS area names. Replacing some names to match place_table names.
    sheet_data = databook[sheet_name].copy()
    #sheet_data = sheet_data[sheet_data['AREA'].str.len() > 0]
    sheet_data['AREA'] = sheet_data['AREA'].str.upper()
    sheet_data.replace(rename_dict, inplace=True)

    #   Just for error checking purposes later. Output contents of unmerged dataframe.
    #sheet_data.to_excel(tempfile_ref, sheet_name, index=None)

    #   Cannot immediately match data_table to match_table because
    #   province names repeat. Pick out province rows in match_table
    #   then combine separately. Do same for HUC, ICC, and Region for convenience.
    prov_df = match_table[match_table['CATEGORY'] == 'Province']
    prov_df = prov_df.join(sheet_data.set_index('AREA'), on='Province')

    city_df = match_table[match_table['CATEGORY'].isin(['HUC', 'ICC'])]
    city_df = city_df.join(sheet_data.set_index('AREA'), on='ALT_LOCAL')

    region_df = match_table[match_table['CATEGORY'] == 'Region']
    region_df = region_df.join(sheet_data.set_index('AREA'), on='ALT_LOCAL')

    #   Add sheet columns to the output workbook and initialize to nan
    match_table[sheet_data.columns] = np.nan

    match_table.update(prov_df)
    match_table.update(city_df)
    match_table.update(region_df)
    print("Writing: " + sheet_name)
    match_table.to_excel(outfile_ref, sheet_name, index=None)

print("Closing")
outfile_ref.save()
outfile_ref.close()
#tempfile_ref.save()
#tempfile_ref.close()
#   ----------
#   ----------
#   ----------
