import requests
from bs4 import BeautifulSoup
import re
import reading_words_form_PDF as Rpdf

#extracting the city names from the URL

def creatingSheetsByName(url):
    city_word = url.split('/')[-1]
    city_name = "".join(re.findall("[a-zA-Z]+", city_word))
    final_name = city_name[:-3]
    return final_name

bs = BeautifulSoup

url = "http://transport.telangana.gov.in/html/reservationnumber.php"

suffix = ".pdf"

link_list = []
response = requests.get(url, stream=True)
soup = bs(response.text)

for link in soup.find_all('a'):    # Finds all links
    if suffix in str(link):   # If the link ends in .pdf    
        link_list.append(link.get('href'))
        
        
for link in link_list:
    city_name = creatingSheetsByName(link)
    link = link.replace('../','http://transport.telangana.gov.in/') 
    pdfUrlResponse = requests.get(link)
    with open('somepdf.pdf','wb') as f:
        f.write(pdfUrlResponse.content)
    Rpdf.writingPDFtoXL(city_name)
    