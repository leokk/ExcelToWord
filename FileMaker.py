import datetime
import os

from Excel import ExcelWrap
from Word import WordWrap


class Filemaker:
    def __init__(self):
        self.cwd = os.getcwd()
        self.file_path = self.cwd + "\\17.xlsx"
        self.excel = ExcelWrap(self.file_path)
        self.word = WordWrap()

    def getHeader(self):
        now = datetime.datetime.now()
        f = open("Header", "r")
        return f.read() + ' ' + str(now.year)

    def totalProfit(self, data):
        summ = 0
        for d in data:
            summ += float(str(d).replace(',', '.'))
        return summ

    def getMain(self, data):
        data_d = {}.fromkeys(data, 0)
        maxx = 0
        max_n = None
        for a in data:
            data_d[a] += 1
            if data_d[a] > maxx:
                maxx = data_d[a]
                max_n = a
        return [max_n, maxx]

    def getProfitByCategories(self, category, profit):
        data_d = {}.fromkeys(category, 0)
        for i in range(self.excel.maxHeight - 1):
            number = round(float(str(profit[i]).replace(',', '.')))
            data_d[category[i]] += number

        return data_d

    def generte(self):
        self.word.addHeader(self.getHeader())
        print(self.excel.getCell(1, 9))

        data = self.getMain(self.excel.getWerticalRange(5, 2, self.excel.maxHeight + 1))

        # word.getStyleList()

        main_ship = "Last year, our company mainly used " + str(data[0]) + " ship mode in count " + str(
            data[1]) + " of " + str(
            self.excel.maxHeight) + " total (" + str(round((data[1] / self.excel.maxHeight) * 100)) + "%). "

        data = self.getMain(self.excel.getWerticalRange(9, 2, self.excel.maxHeight + 1))

        main_interest = "Products were mainly interest to " + str(data[0]) + " customer segment."

        data = self.totalProfit(self.excel.getWerticalRange(6, 2, self.excel.maxHeight + 1))
        total_profit = "Total profit is " + str(round(data, 2)) + "."

        print(main_ship + main_interest + total_profit)

        data = self.getProfitByCategories(self.excel.getWerticalRange(10, 2, self.excel.maxHeight + 1),
                                          self.excel.getWerticalRange(6, 2, self.excel.maxHeight + 1))
        max = 0
        value = None
        for d in range(data.__len__()):
            if float(list(data.values())[d]) > max:
                max = list(data.values())[d]
                value = list(data.keys())[d]

        print(max, value)

        product_max_profit = 'The most profit comes from the sale of ' + value + ' product category (' + str(max) + ')'

        self.word.addParagraph(main_ship + main_interest + total_profit + product_max_profit)

        # word.addInlineExcelChart(file_path, 'D1', 640, 480)

        self.word.addParagraph("Monthly profit is shown in the following chart.")
        self.excel.getChart2()
        self.word.addInlineExcelChart(self.cwd + '\\1.bmp')

        self.word.saveAs(self.cwd + '\\report.docx')

        self.excel.close()
        self.word.close()
