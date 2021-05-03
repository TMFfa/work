import csv
import requests
import json
import pprint

url = 'https://api.lolicon.app/setu/'

r = requests.get(url)
js = json.loads(r.text)
pprint.pprint(js)

with open('csv.csv', 'a+') as f:
    c = csv.writer(f)
    for k, v in js['data'][0].items():
        c.writerow([k, v])

# with open('csv1.csv', 'a+') as f:
#     c = csv.writer(f)
#     c.writerows(js['data'][0])

# with open('csv2.csv', 'a+') as f:
#     c = csv.writer(f)
#     c.writerows(js['data'])
