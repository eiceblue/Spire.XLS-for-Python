from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/templateAz.xlsx"
outputFile = "GetListOfFontsUsed.txt"

#Create a workbook
workbook = Workbook()
#Load a excel document
workbook.LoadFromFile(inputFile)
fonts = []
#Loop all sheets of workbook
for sheet in workbook.Worksheets:
    r = 0
    while r < sheet.Rows.Length:
        for c in sheet.Rows[r].Cells:
            #Get the font of cell and add it to list
            fonts.append(c.Style.Font)
        r += 1
strB = []
for font in fonts:
    strB.append("FontName:{0}; FontSize:{1}".format(font.FontName, font.Size))
AppendAllText(outputFile, strB)
workbook.Dispose()

