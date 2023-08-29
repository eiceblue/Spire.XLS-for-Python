import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/GetFreezePaneRange.xlsx"
outputFile = "GetFreezePaneRange.txt"

#Create a workbook and load a file
wb = Workbook()
wb.LoadFromFile(inputFile)
sheet = wb.Worksheets[0]
rowIndex = None
colIndex = None
#The row and column index of the frozen pane is passed through the out parameter. 
#If it returns to 0, it means that it is not frozen
indexs = sheet.GetFreezePanes()
colIndex = indexs[1]
rowIndex = indexs[0]
r = "Row index: " + str(rowIndex) + ", column index: " + str(colIndex)
#Save the document and launch it
File.AppendAllText(outputFile, r)
wb.Dispose()

