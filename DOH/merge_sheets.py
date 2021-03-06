import pandas as pd
import numpy as np 

#   Merges worksheets by naive concatenate (by row or column).
#   Accomplish vertical merge first before horizontal merge
#   because AREA column is deleted during horizontal merge.

datafile = "Morbidity 2018.xlsx"
outfile = "OUTPUT_" + datafile

print("Running merge...")

lookup_df = pd.read_excel("merge_lookuptable.xlsx", "Sheet1")
data_ref = pd.ExcelFile(datafile)
alldata_df = pd.read_excel(data_ref, data_ref.sheet_names, header=None)

print("Finished loading excel files...")

for row in lookup_df.itertuples(index=False):
    print(row)

    df_src = alldata_df[row.SOURCE]
    df_dest = alldata_df[row.DEST]
    
    axis = 0
    if row.DIRECTION == 'HORIZON':
        df_src.drop([0], axis=1, inplace=True)
        axis = 1

    df_temp = pd.concat([df_dest, df_src], axis=axis, ignore_index=True)

    alldata_df[row.DEST] = df_temp

    del alldata_df[row.SOURCE]

out_ref = pd.ExcelWriter(outfile)

for i, key in enumerate(alldata_df):
    print("Writing " + key)
    #alldata_df[key].reset_index(drop=True, inplace=True)
    #alldata_df[key].loc[-1] = [key]*len(alldata_df[key].columns)
    #alldata_df[key].index = alldata_df[key].index + 1
    #alldata_df[key] = alldata_df[key].sort_index()
    alldata_df[key].to_excel(out_ref, sheet_name=key, index=None, header=None)

print("Saving...")
out_ref.save()
out_ref.close()