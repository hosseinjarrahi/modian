import os
import pandas as pd

# Define the directory path where Excel files are located
dir_path = './files'

# Create an empty list to store data frames
dfs = []

# Loop through each file in the directory
for file_name in os.listdir(dir_path):
    # Check if the file is an Excel file
    df = pd.read_excel(os.path.join(dir_path, file_name))
    
    # Filter rows based on the number of values
    # df = df[df.apply(lambda row: len(row.dropna()) >= 7, axis=1)]
    
    # Append the filtered data frame to the list
    dfs.append(df)

# Create a writer object to write the data frames to an Excel file
writer = pd.ExcelWriter('./output/output.xlsx', engine='xlsxwriter')

# Loop through each data frame and write it to a separate sheet in the Excel file
for i, df in enumerate(dfs):
    sheet_name = 'Sheet{}'.format(i+1)
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer._save()