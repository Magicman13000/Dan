import pandas as pd
import requests
import json

# Load spreadsheet
xl = pd.ExcelFile('test.xlsx')

# Load a sheet into a DataFrame
df = xl.parse('Sheet1')

# Iterate over rows in the DataFrame
for index, row in df.iterrows():
    # Access data using column names
    print(row['ColumnName'])

# Here 'data' is the information you want to send to API
data = {"key": "value"}  # Replace with your data
response = requests.post('https://api.example.com', json=data)

# Make sure the request was successful
if response.status_code == 200:
    json_data = response.json()
    print(json_data)

# Assuming 'json_data' is the data you received from the API
# Convert JSON data to DataFrame
df = pd.DataFrame(json_data)

# Write DataFrame back to Excel
df.to_excel('output.xlsx', index=False)
