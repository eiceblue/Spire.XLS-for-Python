import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


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
File.AppendAllText(outputFile, strB)
workbook.Dispose()

