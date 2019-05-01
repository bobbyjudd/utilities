import urllib.request
from urllib.error import HTTPError
import csv
import random
from bs4 import BeautifulSoup
from googlesearch import search

req = urllib.request.urlopen("https://www.xponential.org/xponential2019/Public/exhibitors.aspx")
page = req.read()
soup = BeautifulSoup(page, 'html.parser')

# Gather list of names from site
ex_list = soup.find_all('a', {'class' : 'exhibitorName'})


# Write each name to the first column
with open('result.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    for link in ex_list:
        name = link.text
        # Google search company
        try:
            for website in search(name, stop=1, pause=1.0*random.randint(2, 7)):
                writer.writerow([name, website])
                print([name, website])
        except HTTPError as e:
            print(e.headers)
            print(e.read())
            print("DIDN'T FINISH!!!!")
            break


print("FINISHED!!!!")
