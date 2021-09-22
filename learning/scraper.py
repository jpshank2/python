#! python3

import requests, re, csv
from bs4 import BeautifulSoup
from datetime import datetime

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}

res = requests.get("https://www.bmss.com/about/our-people/", headers=headers).text

soup = BeautifulSoup(res, "html.parser")

htmlnames = soup.find_all("a", attrs={"class": "t-blue heavy name"})
htmlroles = soup.find_all("span", attrs={"class": "role"})
htmlemail = soup.find_all("a", attrs={"class": "t-green email"})
people = []
roleIndex = 0
emailIndex = 0
peopleIndex = 0

for i in htmlnames:
    x = re.search(r"\B>\w+\s+\w+\s*\w*", str(i))
    people.append({"name": x.group()[1:]})
    
for i in htmlroles:
    x = re.search(r"\B>\w+\s*\w*", str(i))
    people[roleIndex]["role"] = x.group()[1:]
    roleIndex += 1
    #role.append(x.group()[1:])
    #print(x)

for i in htmlemail:
    x = re.search(r"\B</span>\w+@\w+\.\w+", str(i))
    people[emailIndex]["email"] = x.group()[7:]
    emailIndex += 1

with open("C:\\users\\jeremyshank\\desktop\\bmsspeople.csv", "a") as csv_file:
    for i in people:
        writer = csv.writer(csv_file)
        writer.writerow([people[peopleIndex]["name"], people[peopleIndex]["role"], people[peopleIndex]["email"], datetime.now()])
        peopleIndex += 1