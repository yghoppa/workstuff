import pandas as pd

#   This file reads a pdf-to-excel file of the FHSIS and writes the corresponding
#   values according to the place name. Works only for one (1) table.

#   Filename of pdf-to-excel file.
datafilename = "datafile_input.xlsx"

#   Loading entire place_table into a DataFrame.
#   Add new columns for variables/indicators then write to an out file.
output_table = pd.read_excel("place_table.xlsx", "Sheet1")

#   Loading entire data file (pdf-to-excel file of FHSIS) into a DataFrame.
#   Match row place name to output_table (place_table) then add data.
data_table = pd.read_excel(datafilename, "Table 1")

#   Selecting "province" rows, because province name repeats in place_table.
#   Will do the same with HUCs, etc. Set index to province name and use update().
#   Store copy of original index.
prov_df = output_table.loc[output_table['CATEGORY'] == 'Province']
temp_df = pd.DataFrame(index=prov_df.index.copy())
prov_df.set_index('Province', inplace=True)

print(prov_df.index)

prov_df.index = temp_df.index

print(prov_df.index)


#with pd.option_context('display.max_rows', None, 'display.max_columns', None):
#    print(temp_df.index)

#data_table.set_index('AREA')