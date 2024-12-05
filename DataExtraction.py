import requests
from bs4 import BeautifulSoup
import openpyxl
import os  # For folder creation

def extract_article_text(url, url_id, output_folder="extracted_texts"):

    try:
        os.makedirs(output_folder, exist_ok=True)
    except OSError as e:
        print(f"Error creating output folder: {e}")
        return  # Exit function if folder creation fails

    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')

        title = soup.title.string.strip() if soup.title else ""
        article_text = soup.get_text(strip=True)

        # Construct the full path with the folder name
        filename = os.path.join(output_folder, f"{url_id}.txt")

        try:
            with open(filename, "w", encoding='utf-8') as file:
                file.write(f"{title}\n\n{article_text}")
        except OSError as e:
            print(f"Error writing to file '{filename}': {e}")
    else:
        print(f"Error: Failed to download the page. Status code: {response.status_code}")


# Read URLs from Input.xlsx
try:
    wb_in = openpyxl.load_workbook('Input.xlsx')  # Adjust path if needed
    sheet_in = wb_in.active
except FileNotFoundError:
    print("Error: Input.xlsx file not found. Please check the file path.")
    exit()  # Exit program if file not found

# Loop through each URL and extract text
for row in sheet_in.iter_rows(min_row=2):
    url = row[1].value
    url_id = row[0].value
    extract_article_text(url, url_id)