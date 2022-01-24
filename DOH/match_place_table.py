import pandas as pd

#   This file reads a pdf-to-excel file of the FHSIS and writes the
#   corresponding values according to the place name. Works only for
#   one (1) table.

#   Filename of pdf-to-excel file.
datafilename = "datafile_input.xlsx"
outfilename = "output.xlsx"

#   Loading entire place_table into a DataFrame.
#   Add new columns for variables/indicators then write to an out file.
output_table = pd.read_excel("place_table.xlsx", "Sheet1")

#   Loading entire data file (pdf-to-excel file of FHSIS) into a DataFrame.
#   Match row place name to output_table (place_table) then add data.
data_table = pd.read_excel(datafilename, "Table 1")

#   Selecting "province" rows, because province name repeats in place_table.
#   Will do the same with HUCs, etc. Set index to province name and
#   use join(). Store copy of original index and restore later to merge
#   with output_table.
prov_df = output_table.loc[output_table['CATEGORY'] == 'Province']
temp_df = pd.DataFrame(index=prov_df.index.copy())
#prov_df = prov_df.set_index('Province').join(data_table.set_index('AREA'))

[data_table[place == prov_df['Province']] for place in data_table['AREA']]

#prov_df.to_excel(outfilename)