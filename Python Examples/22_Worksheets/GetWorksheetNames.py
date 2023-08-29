import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample3.xlsx"
outputFile = "OutputGetWorksheetNames.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the names of all worksheets
sb = []
for sheet in workbook.Worksheets:
    sb.append(sheet.Name)
#Save the documen
File.AppendAllText(outputFile, sb)
workbook.Dispose()

