import os
import pandas as pd

# Define the directory path where Excel files are located
dir_path = './files'

# Create an empty list to store data frames
dfs = []

# Loop through each file in the directory
for file_name in os.listdir(dir_path):
    # Check if the file is an Excel file
    if file_name.endswith('.csv'):
        # Read the Excel file
        df = pd.read_csv(os.path.join(dir_path, file_name))
        

        
        # Filter rows based on the number of values
        #df = df[df.apply(lambda row: len(row.dropna()) >= 7, axis=1)]
        #df = df.dropna(axis=1, how='all')
        # Append the filtered data frame to the list
        dfs.append(df)


# Create a writer object to write the data frames to an Excel file
writer = pd.ExcelWriter('./output/output_sh.xlsx', engine='xlsxwriter')

# Loop through each data frame and write it to a separate sheet in the Excel file
for i, df in enumerate(dfs):
    sheet_name = 'Sheet{}'.format(i+1)
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer._save()


file_path = './output/output_sh.xlsx'

dfs = pd.read_excel(file_path, sheet_name=None)

# concatenate the dataframes into a single dataframe
result_df = pd.concat(dfs.values(), axis=0, ignore_index=True)
df_sadad = result_df
print(df_sadad)
sadad = list(df_sadad[df_sadad.columns[-10]])

df_bandar = pd.read_excel('./files/یکم الی سی و یک فروردین1402.xls', header=7)
df_bandar = df_bandar[df_bandar.apply(lambda row: len(row.dropna()) >= 7, axis=1)]
df_bandar = df_bandar.dropna(axis=1, how='all')
df_bandar

df_bandar[df_bandar.columns[7]] = df_bandar[df_bandar.columns[7]].astype('int64')

bandar  =  list(df_bandar[df_bandar.columns[7]])

text = 'شماره شبا'
selected_columns = result_df.loc[:, result_df.columns.str.contains(text)]

set1 = set(sadad)
set2 = set(bandar)

diff1 = set1.difference(set2)   # elements in set1 but not in set2
diff2 = set2.difference(set1)   # elements in set2 but not in set1

# display the results
print('Elements in sadad but not in bandar:', diff1)
print('Elements in bandar but not in sadad:', diff2)




diff_sadad = df_sadad[~df_sadad[df_sadad.columns[-10]].isin(df_bandar[df_bandar.columns[7]])]
print("Vsadadlues in df_sadad thsadadt sadadre not in df_bandar:")
print(diff_sadad)


diff_bandar = df_bandar[~df_bandar[df_bandar.columns[7]].isin(df_sadad[df_sadad.columns[-10]])]
print("Vsadadlues in df_bandar thsadadt sadadre not in df_sadad:")
print(diff_bandar)


# create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('./output/diff.xlsx', engine='xlsxwriter')

# write the difference between df_A and df_B to separate sheets in the Excel file
diff_sadad.to_excel(writer, sheet_name='Difference in Sadad')
diff_bandar.to_excel(writer, sheet_name='Difference in Bandar')

# save the Excel file
writer._save()