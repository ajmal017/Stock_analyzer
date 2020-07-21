import requests
import time
from openpyxl import load_workbook

url = "https://stock-google-news.p.rapidapi.com/v1/search"
stock = "AIG"  # stock name
time_frame = "1d"  # time frame

start_time = time.time()
wb = load_workbook("wordslist.xlsx")
negative = wb['Negative']
positive = wb['Positive']
nColumn = negative['A']
pColumn = positive['A']
negativeWL = [nColumn[x].value for x in range(len(nColumn))]
positiveWL = [pColumn[x].value for x in range(len(pColumn))]
querystring = {"when": time_frame, "lang": "en", "country": "US", "ticker": stock}
nro = 0
pNum = 0
nNum = 0
headers = {
    'x-rapidapi-host': "stock-google-news.p.rapidapi.com",
    'x-rapidapi-key': "0e919532damsh36621d13b55468ap12a82cjsnf41c6f3cc086"
}

response = requests.request("GET", url, headers=headers, params=querystring)
t = open("data.txt", "w", encoding='utf-8-sig')
t.write(f"Stock: {stock}\n"f"Timeframe: {time_frame}\n\n")
for (v) in response.json()['articles']:
    nro += 1
    title = ("title: " + str(v['title']))
    for i in positiveWL:
        if i.lower().upper().capitalize() in title:
            pNum += 1
    for i in negativeWL:
        if i.lower().upper().capitalize() in title:
            nNum += 1
    timestamp = ("timestamp: " + str(v['published']))
    source = ("source: " + str(v['link']))
    t.write(f"{nro}.\n{title}\n{timestamp}\n{source}\n\n\n")
t.close()

try:
    total = ((pNum - nNum) / nro) * 100
    h = total
except ZeroDivisionError:
    h = "Summary: 0"

print("[STOCK ANALYSIS]")
print(f"Runtime: {time.time() - start_time}")
print(f"Stock: {stock}")
print(f"Time frame: {time_frame}")
print(f"Articles found: {nro}")
print(f"Positive articles: {pNum}")
print(f"Negative articles: {nNum}")
print('weight: {:.2f}'.format(float(h)))
print("Output file: data.txt")
