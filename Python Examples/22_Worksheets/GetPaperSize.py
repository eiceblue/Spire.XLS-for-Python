import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "GetPaperSize.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#pageInfoList = workbook.GetSplitPageInfo()
sb = []
for sheet in workbook.Worksheets:
    width = sheet.PageSetup.PageWidth
    height = sheet.PageSetup.PageHeight
    sb.append(sheet.Name)
    sb.append("Width: " + str(width) + "\tHeight: " + str(height))
#Save the documen
File.AppendAllText(outputFile, sb)
workbook.Dispose()

