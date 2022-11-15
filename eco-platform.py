from googlesearch import search
import requests
query = "site:epddanmark.dk filetype:pdf"

links = []
i = 0
for j in search(query, tld="co.in", stop=500, pause=1):
    print(str(i) + " " + str(j))
    i += 1
    if j not in links:
        links.append(j)
        print("Saving: " + j)
        response = requests.get(j)
        name = str(j).split('/')
        with open(name[-1], 'wb') as f:
            f.write(response.content)
        with open('links.txt', 'a') as s:
            s.write("\n"+str(j))
    print()