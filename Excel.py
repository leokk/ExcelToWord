import os
from PIL import ImageGrab, Image
import win32com.client as win32

cwd = os.getcwd()


class ExcelWrap:
    maxWidth = None
    maxHeight = None

    def __init__(self, file_path):
        self.Excel = win32.Dispatch("Excel.Application")
        self.workBook = self.Excel.Workbooks.Open(file_path)
        self.workSheet = self.workBook.ActiveSheet
        self.calculateBorders()

    def show(self):
        # convenience when debugging
        self.Excel.Visible = 1

    def close(self):
        self.workBook.Close(True)
        self.Excel.Quit()

    def calculateBorders(self):
        for i in range(1, 100):
            if self.workSheet.Cells(1, i).Text is not "":
                self.maxWidth = i
            else:
                break

        for i in range(1, 1048576):
            if self.workSheet.Cells(i, 1).Text is not "":
                self.maxHeight = i
            else:
                break

    def getWerticalRange(self, w, x1, x2):
        data = []
        for i in range(x1, x2):
            data.append(self.workSheet.Cells(i, w).Text)
        return data

    def getCell(self, x, y):
        return self.workSheet.Cells(x, y).Text

    def getChart(self):
        currentChart = self.workSheet.ChartObjects()
        currentChart.Copy()
        # print(currentChart.Text)
        # currentChart.Chart.Export("chart" + str(0) + ".png")

        image = ImageGrab.grabclipboard()
        print(image)
        # image.save("C:\\Users\\Eugene\\PycharmProjects\\ExcelToWord", 'png')

        # for n, shape in enumerate(self.workSheet.Shapes): shape.Copy() image = ImageGrab.grabclipboard()
        # image.save("C:\\Users\\Eugene\\PycharmProjects\\ExcelToWord", 'png')  # General This code is easy to make
        # mistakes, because the clipboard may have no content, it is best to add try...except... pass


# в якй місяць більше всього купляли
