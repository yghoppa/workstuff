import pandas as pd

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

#   Loading entire place_table into a DataFrame.
#   Add new columns for variables/indicators then write to an out file.
output_table = pd.read_excel("place_table.xlsx", "Sheet1")

#   Loading entire data file (pdf-to-excel file of FHSIS) into a DataFrame.
#   Match row place name to output_table (place_table) then add data.
data_table = pd.read_excel(datafilename, "Table 1")
data_table = data_table.set_index('AREA')

#   Selecting "province" rows, because province name repeats in place_table.
#   Will do the same with HUCs, etc. Set index to province name and
#   use join(). Store copy of original index and restore later to merge
#   with output_table.
prov_df = output_table.loc[output_table['CATEGORY'] == 'Province']
prov_df = prov_df.join(data_table, on='Province')
output_table = output_table.join(prov_df, rsuffix='_in')
temp = output_table['ALT_LOCAL']
output_table['ALT_LOCAL'] = temp.str.upper()
output_table.update(data_table.rename(columns={'AREA': 'ALT_LOCAL'}))

output_table.to_excel(outfilename)

print(output_table.columns)