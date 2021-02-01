import urllib.request
import urllib.parse
import re
import openpyxl as xl

url = 'https://www.bata.com.pk/collections/men-casual'
f = urllib.request.urlopen(url)
data = f.read().decode()

results = re.findall(r'data-original="//(.*?)"(?:\n*.*){27}<div class=\"product-info\">\s*<a href=\".*\">\s*<h3>(.*?)</h3>(?:\n*.*){4}Rs.(([\. 0-9,]+))+</span></div>', data)
colorAndSize = re.findall(r'"option1":"([0-9]+\\/[0-9]+)","option2":"([A-Za-z]+)', data)

wb = xl.load_workbook('ShoeDetailsFile.xlsx')
sheet = wb['Sheet1']
rowNum = 1
cur_row = 1
itemNum = 0
detail = 0
item_number = 0

cols = ['A', 'B', 'C', 'D', ]
cols2 = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']


for itemNum in range(len(results)):
    colNum = 0
    for detail in range(len(results[itemNum])): #url, title, price
        sheet[f'{cols[detail]}{itemNum+1}']= results[itemNum][detail]
    for size in range(itemNum*5, itemNum*5+5):
        sheet[f'{cols2[colNum]}{itemNum+1}'] = colorAndSize[size][0]
        colNum += 1
        sheet[f'{cols2[colNum]}{itemNum+1}'] = colorAndSize[size][1]
        colNum += 1

wb.save('ShoeDetailsFile.xlsx')