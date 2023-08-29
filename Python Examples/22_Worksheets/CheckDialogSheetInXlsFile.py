import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CheckDialogSheetInXlsFile.xlsx"
outputFile = "CheckDialogSheetInXlsFile.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
content = []
if sheet.Type == ExcelSheetType.DialogSheet:
    content.append("Worksheet is a Dialog Sheet!")
else:
    content.append("Worksheet is not a Dialog Sheet!")
File.AppendAllText(outputFile, content)
workbook.Dispose()


