import win32com.client as win32
from win32com.client import constants

import os

from Excel import ExcelWrap
from Word import WordWrap, test

import datetime


def getHeader():
    now = datetime.datetime.now()
    f = open("Header", "r")
    return f.read() + ' ' + str(now.year)


def totalProfit(data):
    summ = 0
    for d in data:
        summ += float(str(d).replace(',', '.'))
    return summ


def getMain(data):
    data_d = {}.fromkeys(data, 0)
    maxx = 0
    max_n = None
    for a in data:
        data_d[a] += 1
        if data_d[a] > maxx:
            maxx = data_d[a]
            max_n = a
    return [max_n, maxx]


def getProfitByCategories(category, profit):
    data_d = {}.fromkeys(category, 0)
    for i in range(excel.maxHeight - 1):
        number = round(float(str(profit[i]).replace(',', '.')))
        data_d[category[i]] += number

    return data_d


cwd = os.getcwd()
file_path = cwd + "\\17.xlsx"
excel = ExcelWrap(file_path)
word = WordWrap()
word.addHeader(getHeader())
print(excel.getCell(1, 9))

data = getMain(excel.getWerticalRange(5, 2, excel.maxHeight + 1))

# word.getStyleList()

main_ship = "Last year, our company mainly used " + str(data[0]) + " ship mode in count " + str(data[1]) + " of " + str(
    excel.maxHeight) + " total (" + str(round((data[1] / excel.maxHeight) * 100)) + "%). "

data = getMain(excel.getWerticalRange(9, 2, excel.maxHeight + 1))

main_interest = "Products were mainly interest to " + str(data[0]) + " customer segment."

data = totalProfit(excel.getWerticalRange(6, 2, excel.maxHeight + 1))
total_profit = "Total profit is " + str(round(data, 2)) + "."

print(main_ship + main_interest + total_profit)

data = getProfitByCategories(excel.getWerticalRange(10, 2, excel.maxHeight + 1),
                             excel.getWerticalRange(6, 2, excel.maxHeight + 1))
max = 0
value = None
for d in range(data.__len__()):
    if float(list(data.values())[d]) > max:
        max = list(data.values())[d]
        value = list(data.keys())[d]

print(max, value)

product_max_profit = 'The most profit comes from the sale of ' + value + ' product category (' + str(max) + ')'

word.addParagraph(main_ship + main_interest + total_profit + product_max_profit)

# word.addInlineExcelChart(file_path, 'D1', 640, 480)

word.addParagraph("Monthly profit is shown in the following chart.")
excel.getChart2()
word.addInlineExcelChart(cwd + '\\1.bmp')

word.saveAs(cwd + '\\report.docx')

excel.close()
word.close()
