import os
import requests
from bs4 import BeautifulSoup
import pandas as pd

# Function to scrape data from the webpage
def scrape_sap_security_patches(url):
    # Send a GET request to the URL
    response = requests.get(url)
    
    # Check if the request was successful
    if response.status_code == 200:
        print("Successfully retrieved the webpage.")
        # Parse the HTML content of the webpage
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find the table containing the security patches
        table = soup.find('table')
        if table:
            print("Table found on the webpage.")
        else:
            print("No table found on the webpage.")
            return None
        
        # Extract data from the table
        rows = table.find_all('tr')
        data = []
        for row in rows[1:]:  # Skip header row
            cols = row.find_all('td')
            cols = [col.text.strip() for col in cols]
            link = row.find('a')['href'] if row.find('a') else ''
            cols[0] = (cols[0], link)  # Include the hyperlink in the first column
            data.append(cols)
        
        print(f"Extracted {len(data)} rows of data.")
        return data
    else:
        print(f"Failed to retrieve data from the webpage. Status code: {response.status_code}")
        return None

# Function to save data to Excel sheet with hyperlinks
def save_to_excel(data, filename, columns):
    # Create a pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
    # Convert the data to a DataFrame
    df = pd.DataFrame(data, columns=columns)
    
    # Write the DataFrame to the Excel file
    df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1, header=False)
    
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Write the column headers with the defined format.
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
    for col_num, value in enumerate(columns):
        worksheet.write(0, col_num, value, header_format)
    
    # Write the data including hyperlinks
    for row_num, row in enumerate(data, start=1):
        for col_num, cell in enumerate(row):
            if col_num == 0:  # The first column contains hyperlinks
                text, link = cell
                worksheet.write_url(row_num, col_num, link, string=text)
            else:
                worksheet.write(row_num, col_num, cell)
    
    # Close the Pandas Excel writer and output the Excel file.
    writer.close()
    print(f"Data saved to {filename}.")

# Example usage
if __name__ == "__main__":
    # Provide the URL of the webpage containing SAP security patches
    url = "https://support.sap.com/en/my-support/knowledge-base/security-notes-news/june-2024.html"
    
    # Scrape data from the webpage
    patches_data = scrape_sap_security_patches(url)
    
    if patches_data:
        # Print current working directory
        print("Current working directory:", os.getcwd())
        
        # Define the correct columns
        columns = ['Patch Number', 'Description', 'Severity', 'CVSS']  # Adjust columns as needed
        
        # Save data to an Excel sheet
        save_to_excel(patches_data, "sap_security_patches.xlsx", columns)
    else:
        print("No data to save.")
