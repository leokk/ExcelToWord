from win32com.client import Dispatch

# cwd = os.getcwd()
# file_path = cwd + "\\16.xlsx"
#
# MYDIR = os.getcwd()


def simpleDemo():
    myWord = Dispatch('Word.Application')
    myWord.Visible = 1

    myDoc = myWord.Documents.Add()
    myRange = myDoc.Range(0, 0)
    myRange.InsertBefore('Hello from Python!')

    # myDoc.SaveAs(MYDIR + '\\python01.doc')
    # myDoc.PrintOut()
    # myDoc.Close()


class WordWrap:
    """Wrapper aroud Word 8 documents to make them easy to build.
    Has variables for the Applications, Document and Selection;
    most methods add things at the end of the document"""

    def __init__(self, templatefile=None):
        self.wordApp = Dispatch('Word.Application')
        if templatefile is None:
            self.wordDoc = self.wordApp.Documents.Add()
        else:
            self.wordDoc = self.wordApp.Documents.Add(Template=templatefile)
        # set up the selection
        self.wordDoc.Range(0, 0).Select()
        self.wordSel = self.wordApp.Selection

    def show(self):
        # convenience when debugging
        self.wordApp.Visible = 1

    def getStyleList(self):
        # returns a dictionary of the styles in a document
        self.styles = []
        stylecount = self.wordDoc.Styles.Count
        for i in range(1, stylecount + 1):
            styleObject = self.wordDoc.Styles(i)
            self.styles.append(styleObject.NameLocal)

    def saveAs(self, filename):
        self.wordDoc.SaveAs(filename)

    def close(self):
        self.wordDoc.Close(True)
        self.wordApp.Quit()

    def printout(self):
        self.wordDoc.PrintOut()

    def selectEnd(self):
        # ensures insertion point is at the end of the document
        self.wordSel.Collapse(0)
        # 0 is the constant wdCollapseEnd; don't weant to depend
        # on makepy support.

    def addText(self, text):
        self.wordSel.InsertAfter(text)
        self.selectEnd()

    def addStyledPara(self, text, stylename):
        if text[-1] != '\n':
            text = text + '\n'

        self.wordSel.InsertAfter(text)
        self.wordSel.Style = stylename
        f = self.wordSel.Range
        self.wordSel.Range.Font.Position = 50
        self.wordSel.Range.Font.Name = 'Times New Roman'

        self.selectEnd()

    def addParagraph(self, text):
        if text[-1] != '\n':
            text = text + '\n'

        self.wordSel.InsertAfter(text)
        self.wordSel.Style = 'По ширині'
        f = self.wordSel.Range
        # self.wordSel.Range.Font.Position = 10
        self.wordSel.Range.Font.Name = 'Times New Roman'
        self.wordSel.Range.Font.Size = 14

        f = self.wordDoc.Range
        self.selectEnd()

    def getStyleList(self):
        for word in self.wordDoc.Styles:
            print(word)

    def addHeader(self, text):
        if text[-1] != '\n':
            text = text + '\n'

        self.wordSel.InsertAfter(text)

        # self.wordSel.Style = 'Цитата'
        self.wordSel.Style = 'Цитата 2'
        f = self.wordSel.Range
        self.wordSel.Range.Font.Position = 10
        self.wordSel.Range.Font.Name = 'Times New Roman'
        self.wordSel.Range.Font.Size = 14
        self.wordSel.Range.Font.Bold = True
        self.wordSel.Range.Font.Italic = False

        f = self.wordDoc.Range
        self.selectEnd()

    def addTable(self, table, styleid=None):
        # Takes a 'list of lists' of data.
        # first we format the text.  You might want to preformat
        # numbers with the right decimal places etc. first.
        textlines = []
        for row in table:
            textrow = map(str, row)  # convert to strings
            textline = str.join(textrow, '\t')
            textlines.append(textline)
        text = str.join(textlines, '\n')

        # add the text, which remains selected
        self.wordSel.InsertAfter(text)

        # convert to a table
        wordTable = self.wordSel.ConvertToTable(Separator='\t')
        # and format
        if styleid:
            wordTable.AutoFormat(Format=styleid)

        self.selectEnd()

    def addInlineExcelChart(self, filename, height=216, width=432):
        # adds a chart inline within the text, caption below.

        # add an InlineShape to the InlineShapes collection
        # - could appear anywhere
        shape = self.wordDoc.InlineShapes.AddPicture(FileName=filename)
        # set height and width in points
        s = self.wordDoc.InlineShapes
        shape.Height = height
        shape.Width = width

        shape.Range.Cut()

        self.wordSel.InsertAfter('chart will replace this')
        self.wordSel.Range.Paste()  # goes in selection

# def randomText():
#     # this may or may not be appropriate in your company
#     RANDOMWORDS = ['strategic', 'direction', 'proactive',
#                    'reengineering', 'forecast', 'resources',
#                    'forward-thinking', 'profit', 'growth', 'doubletalk',
#                    'venture capital', 'IPO']
#
#     sentences = 5
#     output = ""
#     for sentenceno in range(randint(1, 5)):
#         output = output + 'Blah'
#         for wordno in range(randint(10, 25)):
#             if randint(0, 4) == 0:
#                 word = choice(RANDOMWORDS)
#             else:
#                 word = 'blah'
#             output = output + ' ' + word
#         output = output + '.'
#     return output
#
#
# def randomData():
#     months = str.split('Jan Feb Mar Apr')
#     data = []
#     data.append([''] + months)
#     for category in ['Widgets', 'Consulting', 'Royalties']:
#         row = [category]
#         for i in range(4):
#             row.append(randint(10000, 30000) * 0.01)
#         data.append(row)
#     return data
#
#
# def test():
#     outfilename = MYDIR + '\\pythonics_mgt_accounts.doc'
#
#     w = WordWrap()
#     w.show()
#     w.addStyledPara('Accounts for April', 'Назва')
#     w.addHeader("asdasdasdad")
#     # first some text
#     w.addStyledPara("Chairman's Introduction", 'Заголовок 2')
#     w.addStyledPara(randomText(), 'Звичайний')
#
#     # now a table sections
#     w.addStyledPara("Sales Figures for Year To Date", 'Заголовок 1')
#     data = randomData()
#     # w.addTable(data, 37)  # style wdTableStyleProfessional
#     w.addText('\n\n')
#
#     # finally a chart, on the first page of a ready-made spreadsheet
#     w.addStyledPara("Cash Flow Projections", 'Заголовок 1')
#     # w.addInlineExcelChart(MYDIR + '\\wordchart.xls', 'Проста таблиця 2')
#
#     w.saveAs(outfilename)
#     print('saved in', outfilename)
