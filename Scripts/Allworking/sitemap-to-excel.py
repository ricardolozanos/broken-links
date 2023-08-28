

import pandas as pd
import re
import os

# Read content from the text file
with open('sitemap.txt', 'r') as file:
    content = file.read()

# Define a regular expression pattern to match links and dates
pattern = r'(.+?)\s(\d{4}-\d{2}-\d{2})'

# Find all matches of the pattern in the content
matches = re.findall(pattern, content)

# Initialize lists to store links and dates
links = []
dates = []

# Iterate through the matches and extract links and dates
for match in matches:
    link, date = match
    links.append(link)
    dates.append(date)

# Extract folder_name from the first link
first_link = links[0]
folder_name = first_link.split('/')[3]  # Assuming folder_name is the fourth element

# Get the current working directory (where the script is located)
script_directory = os.getcwd()

# Create 'Documents' folder on the current script path if it doesn't exist
documents_folder = os.path.join(script_directory, 'Documents')
os.makedirs(documents_folder, exist_ok=True)

# Create 'Excel_Files' folder inside 'Documents' folder if it doesn't exist
excel_files_folder = os.path.join(documents_folder, 'Excel_Files')
os.makedirs(excel_files_folder, exist_ok=True)

# Create 'TxtFiles' folder inside 'Documents' folder if it doesn't exist
txt_files_folder = os.path.join(documents_folder, 'TxtFiles')
os.makedirs(txt_files_folder, exist_ok=True)

# Save the TXT content to a file
txt_filename = os.path.join(txt_files_folder, f'{folder_name}.txt')
with open(txt_filename, 'w') as txt_file:
    txt_file.write(content)

# Create a DataFrame
data = {'Link': links, 'Template': 'No Carousel', 'Date': dates}
df = pd.DataFrame(data)

# Save DataFrame to an Excel file
excel_filename = os.path.join(excel_files_folder, f'{folder_name}.xlsx')
df.to_excel(excel_filename, index=False)

print(f'Excel file "{excel_filename}" created successfully.')

