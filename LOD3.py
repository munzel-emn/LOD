import pandas as pd
from datetime import datetime, timedelta

# Load the Excel file
file_name = input("Enter your first file name (with extension): ")
engine = input("Enter the engine ('openpyxl' or 'xlrd'): ")  # Specify the engine here
df = pd.read_excel(file_name, engine=engine)
print(df.columns)
print(df.head())

# Specify the columns you want to keep
columns_to_keep = ['Container Number', 'LOD Date', 'Empty', 'Stowage Position', 'Consignor Code',]
df = df[columns_to_keep]

# Create new Columns and set value
df['Type'] = 'LOD'
df['Depot'] = input('Enter depot name: ')

# Convert the 'Stowage Position' column to string
df['Stowage Position'] = df['Stowage Position'].astype(str)

# Add "Stow Position-0" followed by the digits
df['Stowage Position'] = 'Stow Position-0' + df['Stowage Position'].str.zfill(6)

# Prompt the user for a date and time
input_date_time_str = input("Enter a date and time (e.g., '2023-09-24 12:30:00'): ")
input_date_time = datetime.strptime(input_date_time_str, '%Y-%m-%d %H:%M:%S')

# Loop through the dataframe and subtract five minutes from each subsequent row
for index, row in df.iterrows():
    df.at[index, 'LOD Date'] = input_date_time.strftime('%Y-%m-%d %H:%M:%S')

# Loop through the dataframe and update 'Consignor Code' based on the condition
for index, row in df.iterrows():
    container_number = row['Container Number']
    empty = row['Empty']

    if not isinstance(container_number, str):
        print(f"Row {index}: Unexpected data type for 'Container Number': {type(container_number)}")
        continue  # Skip this row

    if container_number.startswith(('CELU', 'CELR', 'TRIU', 'TKKU')) and empty == 'Y':
        df.at[index, 'Consignor Code'] = 'CELGEN'

# Set 'Principal' column based on 'Container Number' and 'Consignor'
def get_principal(row):
    container_number = row['Container Number']
    consignor_code = row['Consignor Code']

    if container_number.startswith(('CELU', 'CELR', 'TRIU', 'TKKU', 'LSSR')):
        return 'CELGEN'
    elif consignor_code in ('SWIMTS', 'SWIRES', 'SWIIMP'):
        return 'SWIRES'
    elif consignor_code == 'CARPSA':
        return 'CARPSA'
    elif consignor_code == 'PUMAEN':
        return 'PUMAEN'
    else:
        return ''

df['Principal'] = df.apply(get_principal, axis=1)

# Rearrange columns
reordered_columns = [
    'Container Number', 'Stowage Position', 'Consignor Code', 'Type',
    'LOD Date', 'Depot', 'Empty', 'Principal'
]
df = df[reordered_columns]

# Save the DataFrame as a CSV file
output_csv_file = input("Enter the output CSV file name: ")
df.to_csv(output_csv_file, index=False)

# Display the current data in the specified columns
print("Current data in specified columns:")
print(df.columns)
print(df.head(10))

print("Updated data saved to the CSV file.")
