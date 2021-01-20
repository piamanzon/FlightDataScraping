import requests, bs4


res = requests.get('https://www.yvr.ca/en/passengers/flights/arriving-flights')
res.raise_for_status()
#noStarchSoup = bs4.BeautifulSoup(res.text, 'html.parser')
noStarchSoup = bs4.BeautifulSoup(res.text, 'html.parser')
table = noStarchSoup.select(405435562)
print(table)

