# from selenium.webdriver.chrome.service import Service
# from selenium import webdriver
#
# options = webdriver.FirefoxOptions()
# # options.add_argument('--headless')
# # options.add_argument('--no-sandbox')
# # options.add_argument('--disable-dev-shm-usage')
# service = Service(executable_path="geckodriver.exe")
# wd = webdriver.Firefox(service=service, options=options)
# URL = "https://www.eco-platform.org/epd-data.html"
# wd.get(URL)
# print(wd)
#

import requests
URL = "https://www.greenbooklive.com/pdfdocs/en15804epd/BREGENEPD000"
name = "BREGENEPD000"
i = 1

while i < 500:
    if i <= 9:
        x = "00"+str(i)
    elif i > 9 and i <= 99:
        x = "0"+str(i)
    elif i > 99 and i < 999:
        x = str(i)
    response = requests.get(URL+x+".pdf")
    print(str(i)+ " " + URL+x+".pdf")
    if response.status_code != 404:
        print("Saving: " + str(name+str(i)+'.pdf\n'))
        with open(str(name+str(i)+'.pdf'), 'wb') as f:
            f.write(response.content)
    else:
        print("ERROR 404 at ", i)
    i += 1

