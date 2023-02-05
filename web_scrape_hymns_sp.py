import requests
from bs4 import BeautifulSoup

url = 'https://singpraises.net/charts/cross-reference?a=en_hymns&b=es_hymns&show_both_numbers=true'# XXX:
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

hymn_titles = [h.text.strip() for h in soup.find_all('span', {'class': 'title'})]

sp_hymn_numbers = [h.text for h in soup.find_all('span', {'style':"display: inline-block; width: 2.15em; text-align: right;"})]
en_hymn_numbers = [h.text for h in soup.find_all('span', {'class':'number'})][1:]

print(hymn_titles)
print(sp_hymn_numbers)
print(en_hymn_numbers)

trimmed_en = []
for i,en in enumerate(en_hymn_numbers):
    splength = len(sp_hymn_numbers[i])
    englength = len(en)
    diff = englength - splength
    trimmed_en.append(en[:diff])

print(trimmed_en)

import pandas as pd

df = pd.DataFrame({'spanish_number':sp_hymn_numbers,'english_number':trimmed_en})
print(df)

df.to_csv('cross_ref.csv',index=False)
