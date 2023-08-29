import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


outputFile = "GetDefaultRowAndColumnCount.txt"

#Create a workbook
workbook = Workbook()
#Clear all worksheets
workbook.Worksheets.Clear()
#Create a new worksheet
sheet = workbook.CreateEmptySheet()
sb = []
#Get row and column count
rowCount = sheet.Rows.Length
columnCount = sheet.Columns.Length
sb.append("The default row count is :" + str(rowCount))
sb.append("The default column count is :" + str(columnCount))
File.AppendAllText(outputFile, sb)

