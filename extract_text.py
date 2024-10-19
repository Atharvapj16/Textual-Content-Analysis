import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

def extract_article_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    
    title = soup.find('h1').get_text() if soup.find('h1') else 'No Title'
    
   
    paragraphs = soup.find_all(['p', 'ol', 'li'])
    article_text = ' '.join([element.get_text() for element in paragraphs])
    
    return title + '\n' + article_text

input_file = 'Input.xlsx'
df = pd.read_excel(input_file)
os.makedirs('articles', exist_ok=True)

for index, row in df.iterrows():
    url = row['URL']
    url_id = row['URL_ID']
    article_text = extract_article_text(url)
    
    with open(f'articles/{url_id}.txt', 'w', encoding='utf-8') as file:
        file.write(article_text)