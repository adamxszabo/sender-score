import requests, datetime, openpyxl
from bs4 import BeautifulSoup

#get sender score for the selected IP address
site = requests.get("https://senderscore.org/report/?lookup=YOUR-IP-ADDRESS&authenticated=true", headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.76 Safari/537.36', "Upgrade-Insecure-Requests": "1","DNT": "1","Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8","Accept-Language": "en-US,en;q=0.5","Accept-Encoding": "gzip, deflate"}, verify = False).text
soup = BeautifulSoup(site, "html.parser")
h2 = soup.find("h2", {"class" : "rp-text-color-green-light"})
senderScore = int(h2.text.strip())

#get the current date
timeNow = datetime.datetime.now()
dateToday = timeNow.date()

#open the excel file
excelFile = openpyxl.load_workbook("sender_score.xlsx")
excelSheet = excelFile.worksheets[0]

#add current date to the first column
emptyA = "A" + str(excelSheet.max_row + 1)
excelSheet[emptyA] = dateToday

#add sender score to the second column
emptyB = "B" + str(excelSheet.max_row)
excelSheet[emptyB] = senderScore

#save the excel file
excelFile.save("sender_score.xlsx")
excelFile.close()
