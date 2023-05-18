import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
from fuzzywuzzy import fuzz
import os
import re

# Get the script directory path
script_dir = os.path.dirname(os.path.abspath(__file__))

# Specify the Excel file name
excel_file_name = 'VendorHoldList.xlsx'
sheet_name = 'Vendor List (2023)'

# Construct the Excel file path
excel_file_path = os.path.join(script_dir, excel_file_name)

# Load data from Excel
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Column names in the Excel sheet
name_column = 'Name'
vendor_id_column = 'Vendor ID'
date_checked_column = 'Date Checked'
hold_status_column = 'Hold Status'

# URL of the FMCPA website
fmcpa_url = 'https://fmcpa.cpa.state.tx.us/tpis/servlet/TPISReports'

print("Script running...")
# Create a new DataFrame for search results
result_data = {name_column: [], vendor_id_column: [], hold_status_column: []}

# Iterate over each vendor in the Excel sheet
for index, row in df.iterrows():
    name = row[name_column]
    vendor_id = row[vendor_id_column]

    # Remove special characters from the vendor name
    name = re.sub(r'[^a-zA-Z0-9\s]', '', name)

    # Prepare the payload for the website search
    payload = {
        'reptId': 'wrntHoldSrc',
        'searchString': name,
        'orderBy': 'rel'
    }

    # Send the request to the FMCPA website
    response = requests.get(fmcpa_url, params=payload)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Check if the vendor is on hold
    table = soup.find('table', summary='Vendor name, location and a calculated score of how close each item matches the search terms.')
    if table:
        rows = table.find_all('tr')
        if len(rows) > 1:
            vendor_name = rows[1].find('th', scope='row')
            if vendor_name:
                vendor_name = vendor_name.text.strip()
                vendor_name = re.sub(r'[^a-zA-Z0-9\s]', '', vendor_name)
                relevance = rows[1].find_all('td')[-1].text.strip()
                print(f'Searching Vendor: {name} | Found Vendor: {vendor_name} | Relevance: {relevance}')
                if 'Vendor Hold' in vendor_name:
                    hold_status = 'On Hold'
                elif fuzz.token_sort_ratio(name.lower(), vendor_name.lower()) >= 80:
                    hold_status = 'Likely On Hold'
                else:
                    hold_status = ''
            else:
                print(f'Searching Vendor: {name} | No vendor found in the search result')
                hold_status = ''
        else:
            print(f'Searching Vendor: {name} | No vendor found in the search result')
            hold_status = ''
    else:
        print(f'Searching Vendor: {name} | No vendor found in the search result')
        hold_status = ''

    # Append the vendor details and hold status to the result DataFrame
    result_data[name_column].append(name)
    result_data[vendor_id_column].append(vendor_id)
    result_data[hold_status_column].append(hold_status)

# Create a DataFrame from the result data
result_df = pd.DataFrame(result_data)

# Create a new Excel file for the search results
result_file = os.path.join(script_dir, 'SearchResults.xlsx')
with pd.ExcelWriter(result_file) as writer:
    result_df.to_excel(writer, sheet_name='Search Results', index=False)

print("Script completed")
