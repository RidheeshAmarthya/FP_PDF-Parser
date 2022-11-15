import requests
from bs4 import BeautifulSoup
import time

url = 'https://www.epd-norge.no/epder/?fromUrl=epder%2F&offset704='
i = 2190
count = 810
links = []
while i <= 2190:
    soup = BeautifulSoup(requests.get(url+str(i)).content, "html.parser")
    for div in soup.findAll("div", {"class": "thumbnail"}):
        link = [a.get('href') for a in div.find_all('a', href=True)]
        if link not in links:

            pdf_l = BeautifulSoup(requests.get(link[0]).content,
                                  'html.parser')
            for div in pdf_l.findAll("div", {"class": "data"}):
                pdf = [a.get('href') for a in div.find_all("a", {"data-type": "epd"}, href=True)]
                if pdf:
                    name = str(pdf[0]).split('/')
                    respons1 = requests.get("https:"+pdf[0])
                    with open(name[-1], 'wb') as f:
                        f.write(respons1.content)
                    print("Saved: ", name[-1])

            links.append(link)
            print(count, " ", link)
            with open('links2.txt', 'a') as s:
                s.write(link[0] + ',\n')
            s.close()
            print("Saved")
            count += 1
    i += 30
#73 pages
#825