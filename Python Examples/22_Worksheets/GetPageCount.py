import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "GetPageCount.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
pageInfoList = workbook.GetSplitPageInfo()
sb = []
for i, unusedItem in enumerate(workbook.Worksheets):
    sheetname = workbook.Worksheets[i].Name
    pagecount = pageInfoList[i].Count
    sb.append(sheetname + "'s page count is: " + str(pagecount))
#Save the documen
File.AppendAllText(outputFile, sb)
workbook.Dispose()
