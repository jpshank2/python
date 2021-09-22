import requests, random
import win32com.client as wc

# people = requests.get('http://api.open-notify.org/astros.json')
# people_json = people.json()

# print("Number of poeple in space:", people_json['number'])

# for person in people_json['people']:
#     print(person['name'])


# rhyme = 'jingle'#input()
# parameter = {"rel_rhy":rhyme}

# request = requests.get('https://api.datamuse.com/words', parameter)

# rhyme_json = request.json()
# for i in rhyme_json[0:3]:
#     print(i['word'])

season = random.randint(1, 10)
if season == 1:
    episode = random.randint(1, 7)
elif season == 2:
    episode = random.randint(1, 23)
elif season == 4:
    episode = random.randint(1, 20)
elif season == 5:
    episode = random.randint(1, 29)
elif season == 7:
    episode = random.randint(1, 27)
elif season == 8:
    episode = random.randint(1, 25)
else:
    episode = random.randint(1, 26)

uri = 'https://the-office-api.herokuapp.com/season/' + str(season) + '/episode/' + str(episode)
quotesInJSON = requests.get(uri)
quotes = quotesInJSON.json()
quoteIndex = random.randint(0, len(quotes['data']['quotes']))
quote = quotes['data']['quotes'][quoteIndex][0]

outlook = wc.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'jshank@abacustechnologies.com;cboyd@abacustechnologies.com'
mail.Subject = 'Office Quote of the Day'
mail.Body = quote
mail.Send()