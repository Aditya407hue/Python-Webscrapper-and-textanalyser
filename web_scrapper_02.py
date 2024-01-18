import openpyxl
from bs4 import BeautifulSoup
import requests

#for url of - 14,20,29,43,83
# Function to extract article text from a given URL
def extract_article_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Assuming the title is in an HTML element with the 'title' class
    title = soup.find('h1', class_='tdb-title-text').get_text()
    # Assuming the article text is in an HTML element with the 'article' class

    article_elements = soup.find_all('div', class_='td-post-content')
    article_text = '\n'.join([element.get_text() for element in article_elements]) if article_elements else "Article text not found"

    return title, article_text


# Load URLs from the Excel file
workbook = openpyxl.load_workbook('Input-2.xlsx')
sheet = workbook.active

# Iterate through each row in the Excel file
for row in sheet.iter_rows(min_row=2, values_only=True):
    url_id, url = row

    try:
        # Extract article text
        title, article_text = extract_article_text(url)

        # Save the extracted article to a text file
        with open(f'{url_id}.txt', 'w', encoding='utf-8') as file:
            file.write(f'Title: {title}\n\n')
            file.write(f'Article Text:\n{article_text}\n')

        print(f'Successfully extracted and saved article for {url_id}')

    except Exception as e:
        print(f'Error processing {url_id}: {e}')

# Close the Excel file
workbook.close()
